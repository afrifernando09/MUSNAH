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

# Input API Key di sidebar
with st.sidebar:
    st.header("Pengaturan")
    api_key = st.text_input("Masukkan Gemini API Key", type="password")
    st.markdown("[Dapatkan API Key Gratis di sini](https://aistudio.google.com/app/apikey)")

# Area Upload File
col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("1. Upload MASTER STOK (.xls / .xlsx)", type=["xls", "xlsx"])
with col2:
    image_file = st.file_uploader("2. Upload Gambar Catatan (.jpg / .png)", type=["jpg", "jpeg", "png"])

def calculate_qty(qty_expr):
    if not qty_expr or not str(qty_expr).strip():
        return 0
    try:
        clean_expr = re.sub(r'[^\d+\-*/. ]', '', str(qty_expr))
        return eval(clean_expr) if clean_expr else 0
    except Exception:
        return 0

def cari_kecocokan(plu_tulis, desk_tulis, df_master):
    # Bersihkan spasi dan huruf yang mirip angka pada PLU
    plu_bersih = str(plu_tulis).replace('O', '0').replace('o', '0').replace('D', '0').replace(' ', '').strip()
    
    # Tahap 1: Pencarian PLU Sama Persis
    match_plu = df_master[df_master['PLU'] == plu_bersih]
    if not match_plu.empty:
        return match_plu.iloc[0]['PLU'], match_plu.iloc[0]['DESKRIPSI'], "✅ Cocok (Dari PLU)"
        
    # Tahap 2: Pencarian Sebagian Angka PLU (jika angka agak terpotong)
    if len(plu_bersih) >= 4:
        match_contain = df_master[df_master['PLU'].str.contains(plu_bersih, na=False)]
        if not match_contain.empty:
            return match_contain.iloc[0]['PLU'], match_contain.iloc[0]['DESKRIPSI'], "✅ Cocok (PLU Mirip)"
            
    # Tahap 3: PENCARIAN PINTAR BERDASARKAN NAMA BARANG
    desk_clean = str(desk_tulis).upper()
    # Ambil kata-kata penting yang panjangnya lebih dari 2 huruf (mengabaikan singkatan seperti 'B', 'BO')
    kata_kunci = [k for k in re.findall(r'[A-Z]+', desk_clean) if len(k) > 2]
    
    if kata_kunci:
        def hitung_skor(teks_master):
            teks_master = str(teks_master).upper()
            # Hitung berapa banyak kata kunci dari catatan yang ada di Master Excel
            skor = sum(1 for k in kata_kunci if k in teks_master)
            return skor
            
        skor_semua = df_master['DESKRIPSI'].apply(hitung_skor)
        skor_maksimal = skor_semua.max()
        
        # Jika sistem menemukan kecocokan kata (misal: tulisan "B Bombay" menemukan kata "BOMBAY" di Master)
        if skor_maksimal > 0:
            baris_terbaik = df_master.loc[skor_semua.idxmax()]
            return baris_terbaik['PLU'], baris_terbaik['DESKRIPSI'], "⚠️ Cocok (Dari Nama/Teks)"

    # Tahap 4: Gagal menemukan sama sekali
    return plu_tulis, f"{desk_tulis} (TIDAK DITEMUKAN)", "❌ Tidak Cocok"

if st.button("🚀 Proses Data Sekarang", use_container_width=True):
    if not api_key:
        st.error("⚠️ Masukkan Gemini API Key terlebih dahulu di menu samping!")
    elif not master_file or not image_file:
        st.error("⚠️ Harap upload file Excel dan Gambar terlebih dahulu!")
    else:
        with st.spinner('Sedang membaca gambar dan mencocokkan data... Mohon tunggu...'):
            try:
                # 1. Konfigurasi AI 
                genai.configure(api_key=api_key)
                # Menggunakan model 1.5-flash karena lebih cepat dan sangat stabil
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                # 2. Baca Gambar
                img = Image.open(image_file)
                prompt = """
                Ekstrak data dari gambar catatan ini.
                Format output HARUS berupa JSON array of objects. 
                Setiap object mewakili satu baris data dengan struktur:
                [
                  {
                    "plu_tulis": "KODE PLU (angka di kolom pertama)",
                    "deskripsi": "Deskripsi barang",
                    "qty_expr": "Operasi matematika di kolom terakhir (misal: '395 + 135 + 400', '120')"
                  }
                ]
                Pastikan huruf besar/kecil sesuai gambar. Hanya berikan JSON murni, tanpa teks penjelasan apapun.
                """
                
                response = model.generate_content([img, prompt])
                cleaned_response = response.text.replace("```json", "").replace("```", "").strip()
                extracted_data = json.loads(cleaned_response)
                
                # 3. Baca Excel
                df_master = pd.read_excel(master_file, sheet_name=0)
                df_master['PLU'] = df_master['PLU'].astype(str).str.strip()
                df_master['DESKRIPSI'] = df_master['DESKRIPSI'].fillna('')
                
                # 4. Proses Pencocokan
                results = []
                for item in extracted_data:
                    plu_tulis = str(item.get('plu_tulis', '')).strip()
                    desk_tulis = str(item.get('deskripsi', '')).strip()
                    qty_expr = item.get('qty_expr', '')
                    
                    total_qty = calculate_qty(qty_expr)
                    
                    # PANGGIL FUNGSI PENCOCOKAN PINTAR
                    plu_final, desk_final, status = cari_kecocokan(plu_tulis, desk_tulis, df_master)
                        
                    results.append({
                        'PLU Master': plu_final,
                        'Deskripsi Master': desk_final,
                        'Total Qty': total_qty,
                        'Status': status,
                        'Tulisan Asli (PLU - Nama)': f"{plu_tulis} - {desk_tulis}",
                        'Hitungan Asli': qty_expr
                    })
                
                df_result = pd.DataFrame(results)
                
                # 5. Tampilkan Hasil
                st.success("Berhasil memproses data!")
                st.dataframe(df_result, use_container_width=True)
                
                # 6. Tombol Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Hasil Rekap')
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Hasil Excel",
                    data=output,
                    file_name="Hasil_Rekap_Catatan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
