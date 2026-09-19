import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
import re
from PIL import Image
import io

st.set_page_config(page_title="Rekap Catatan Stok", layout="wide")

st.title("📝 Aplikasi Rekap Catatan Stok Otomatis")
st.write("Upload gambar catatan tangan dan file MASTER STOK Anda di sini.")

with st.sidebar:
    st.header("Pengaturan")
    api_key = st.text_input("Masukkan Gemini API Key", type="password")
    st.markdown("[Dapatkan API Key Gratis di sini](https://aistudio.google.com/app/apikey)")

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("1. Upload MASTER STOK (.xls / .xlsx)", type=["xls", "xlsx"])
with col2:
    image_file = st.file_uploader("2. Upload Gambar Catatan (.jpg / .png)", type=["jpg", "jpeg", "png"])

def calculate_qty(qty_expr):
    if not qty_expr or str(qty_expr).strip() == "0":
        return 0
    try:
        clean_expr = re.sub(r'[^\d+\-*/. ]', '', str(qty_expr))
        return eval(clean_expr) if clean_expr else 0
    except Exception:
        return 0

def clean_deskripsi(text):
    text = str(text).upper().replace('B/O', 'BUAH OLAHAN').replace('BO ', 'BUAH OLAHAN ')
    replace_dict = {'B BOMBAI': 'BAWANG BOMBAY', 'B MERAH': 'BAWANG MERAH', 'B PUTIH': 'BAWANG PUTIH', 'B BIMA': 'BAWANG BIRMA'}
    for k, v in replace_dict.items():
        if text.startswith(k) or k in text:
            text = text.replace(k, v)
    return text

def get_robust_match(plu_tulis, deskripsi_tulis, is_pck, df_master):
    plu_bersih = str(plu_tulis).replace('D', '0').replace('O', '0').strip()
    desk_bersih = clean_deskripsi(deskripsi_tulis)
    pck_units = ['PCS', 'PCK', 'IKT', 'CTN', 'CUP']
    
    def evaluate(df_sub):
        if df_sub.empty: return None
        df_sub = df_sub.copy()
        df_sub['Has_Stock'] = (df_sub['STOK'] != 0).astype(int)
        df_sub['Unit_Match'] = df_sub['UNIT'].isin(pck_units).astype(int) if is_pck else 1
        return df_sub.sort_values(by=['Has_Stock', 'Unit_Match'], ascending=[False, False]).iloc[0]

    # 1. Exact PLU
    exact_plu = df_master[df_master['PLU'] == plu_bersih]
    if not exact_plu.empty:
        return evaluate(exact_plu)
        
    # 2. PLU mirip / terkandung di dalamnya
    if len(plu_bersih) >= 5:
        sim_plu = df_master[df_master['PLU'].str.contains(plu_bersih, na=False)]
        if not sim_plu.empty:
            return evaluate(sim_plu)
            
    # 3. Pencarian Teks Pintar
    keywords = [k for k in re.findall(r'[A-Z0-9]+', desk_bersih) if len(k) > 2]
    if 'BUAH OLAHAN' in desk_bersih:
        keywords = ['BUAH OLAHAN'] + [k for k in keywords if k not in ['BUAH', 'OLAHAN']]
    
    if keywords:
        def calc_score(text):
            text_up = str(text).upper()
            return sum(1 for k in keywords if k in text_up)
        
        df_master['score'] = df_master['DESKRIPSI'].apply(calc_score)
        max_score = df_master['score'].max()
        if max_score > 0:
            return evaluate(df_master[df_master['score'] == max_score])
            
    return None

def apply_grouping(df_results, max_sum=495000):
    df_results['Kelompok'] = 0
    current_group = 1
    current_sum = 0
    for index, row in df_results.iterrows():
        if current_sum + row['Total Rupiah'] > max_sum and current_sum > 0:
            current_group += 1
            current_sum = row['Total Rupiah']
        else:
            current_sum += row['Total Rupiah']
        df_results.at[index, 'Kelompok'] = current_group
    return df_results

# Fungsi Pewarnaan Tabel
def color_groups(row):
    colors = ['#e6f2ff', '#e6ffe6', '#ffffe6', '#ffe6e6', '#f9e6ff', '#e6ffff', '#fff0e6']
    color = colors[(row['Kelompok'] - 1) % len(colors)]
    return [f'background-color: {color}'] * len(row)

if st.button("🚀 Proses Data Sekarang", use_container_width=True):
    if not api_key or not master_file or not image_file:
        st.error("⚠️ Pastikan API Key, File Excel, dan Gambar telah diunggah!")
    else:
        with st.spinner('Memproses data, menghitung stok, dan mewarnai kelompok...'):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-3.5-flash')
                
                img = Image.open(image_file)
                prompt = """
                Ekstrak data catatan. Output berupa JSON array of objects:
                [
                  {
                    "plu_tulis": "KODE PLU (angka)",
                    "deskripsi": "Deskripsi barang",
                    "qty_expr": "Operasi matematika misal '395 + 135' (termasuk kata pcs/pck jika ada)"
                  }
                ]
                Hanya berikan JSON murni.
                """
                response = model.generate_content([img, prompt])
                extracted_data = json.loads(response.text.replace("```json", "").replace("```", "").strip())
                
                df_master = pd.read_excel(master_file, sheet_name=0)
                df_master['PLU'] = df_master['PLU'].astype(str).str.strip()
                df_master['STOK'] = pd.to_numeric(df_master.get('STOK', 0), errors='coerce').fillna(0)
                df_master['ACOST'] = pd.to_numeric(df_master.get('ACOST', 0), errors='coerce').fillna(0)
                df_master['FRAC'] = pd.to_numeric(df_master.get('FRAC', 1), errors='coerce').fillna(1).replace(0, 1)
                
                results = []
                for item in extracted_data:
                    plu = str(item.get('plu_tulis', '')).strip()
                    desk = str(item.get('deskripsi', '')).strip()
                    qty_expr = item.get('qty_expr', '')
                    
                    is_pck = bool(re.search(r'\b(pcs|pck)\b', qty_expr.lower()) or re.search(r'\b(pcs|pck)\b', desk.lower()))
                    total_qty = calculate_qty(qty_expr)
                    
                    match = get_robust_match(plu, desk, is_pck, df_master)
                    
                    if match is not None:
                        acost = float(match['ACOST'])
                        frac = float(match['FRAC'])
                        total_rupiah = (total_qty * acost) / frac
                        
                        results.append({
                            'Kelompok': 0,
                            'PLU Tulis': plu,
                            'Deskripsi Asli': desk,
                            'PLU Master': match['PLU'],
                            'Deskripsi Master': match['DESKRIPSI'],
                            'Qty': total_qty,
                            'Unit': match['UNIT'],
                            'Stok System': match['STOK'],
                            'Total Rupiah': round(total_rupiah, 2)
                        })
                    else:
                        results.append({
                            'Kelompok': 0, 'PLU Tulis': plu, 'Deskripsi Asli': desk, 
                            'PLU Master': '-', 'Deskripsi Master': 'TIDAK DITEMUKAN', 
                            'Qty': total_qty, 'Unit': '-', 'Stok System': 0, 'Total Rupiah': 0.0
                        })
                
                df_result = pd.DataFrame(results)
                df_result = apply_grouping(df_result, max_sum=495000)
                
                st.success("Berhasil memproses data!")
                st.dataframe(df_result.style.apply(color_groups, axis=1).format({'Total Rupiah': "Rp {:,.2f}"}), use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Hasil Rekap')
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Hasil Excel",
                    data=output,
                    file_name="Hasil_Rekap_Dikelompokkan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
