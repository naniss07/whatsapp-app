import io
import os
import tempfile
from datetime import datetime

import streamlit as st

# Mevcut iş mantığı `test.py` içinde; aynen reuse ediyoruz.
from test import process_files, write_excel


st.set_page_config(page_title="WhatsApp Metin → Excel", layout="centered")

st.title("WhatsApp metinlerinden Excel özet çıkarma")

st.write(
    "WhatsApp export `.txt` dosyalarını yükle, istersen tarih aralığı ver, Excel çıktısını indir."
)

uploaded = st.file_uploader(
    "WhatsApp export .txt dosyaları",
    type=["txt"],
    accept_multiple_files=True,
)

col1, col2 = st.columns(2)
with col1:
    from_str = st.text_input("Başlangıç (DD.MM.YYYY) - opsiyonel", value="")
with col2:
    to_str = st.text_input("Bitiş (DD.MM.YYYY) - opsiyonel", value="")

out_name = st.text_input("Çıktı dosya adı", value="ariza_takip.xlsx")

run = st.button("Excel üret")

if run:
    if not uploaded:
        st.error("En az 1 adet `.txt` dosyası yüklemen gerekiyor.")
        st.stop()

    from_arg = from_str.strip() or None
    to_arg = to_str.strip() or None
    if (from_arg and not to_arg) or (to_arg and not from_arg):
        st.error("Tarih filtresi için hem başlangıç hem bitiş girilmeli (ya da ikisi de boş).")
        st.stop()

    with st.spinner("İşleniyor..."):
        temp_dir = tempfile.mkdtemp(prefix="wa_txt_")
        paths: list[str] = []
        try:
            for f in uploaded:
                safe_name = os.path.basename(f.name) or "whatsapp.txt"
                p = os.path.join(temp_dir, safe_name)
                with open(p, "wb") as w:
                    w.write(f.getbuffer())
                paths.append(p)

            rows = process_files(paths, from_arg, to_arg)

            out_xlsx = os.path.join(temp_dir, "out.xlsx")
            write_excel(rows, out_xlsx)

            with open(out_xlsx, "rb") as r:
                data = r.read()

        finally:
            # Streamlit dosyayı RAM'den indirtiyor; temp klasörü silmek şart değil.
            # Windows'ta bazı durumlarda dosya kilidi kalabiliyor, o yüzden sessiz geçiyoruz.
            pass

    st.success(f"{len(rows)} kayıt bulundu.")
    st.download_button(
        label="Excel indir",
        data=io.BytesIO(data),
        file_name=out_name or f"ariza_takip_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
