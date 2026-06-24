import streamlit as st
import pandas as pd
from pdf2docx import Converter
import PyPDF2
from docx import Document
import io
import os

# إعدادات الواجهة (تعديل العنوان)
st.set_page_config(page_title="Fily - المحول العالمي", layout="wide")

# العنوان الرئيسي في الصفحة
st.title("🚀 Fily - المحول العالمي")

# 2. قاموس الـ 20 لغة (مكتمل بالكامل لتغيير لغة الموقع)
langs_dict = {
    "العربية": {"title": "محول المستندات والجداول", "service": "اختر التحويل:", "up": "أحضر الملف هنا", "conv": "تحويل الملف", "down": "تنزيل الملف الجاهز", "wait": "جاري التحويل...", "success": "تم التحويل بنجاح!", "error": "حدث خطأ في الملف"},
    "English": {"title": "Docs & Tables Converter", "service": "Choose Conversion:", "up": "Upload File Here", "conv": "Convert File", "down": "Download Ready File", "wait": "Converting...", "success": "Converted Successfully!", "error": "File Error"},
    "Français": {"title": "Convertisseur de Documents", "service": "Choisir la conversion:", "up": "Déposer le fichier ici", "conv": "Convertir le fichier", "down": "Télécharger le fichier prêt", "wait": "Conversion en cours...", "success": "Converti avec succès!", "error": "Erreur de fichier"},
    "Español": {"title": "Convertidor de Documentos", "service": "Elegir conversión:", "up": "Subir archivo aquí", "conv": "Convertir archivo", "down": "Descargar archivo listo", "wait": "Convirtiendo...", "success": "¡Convertido con éxito!", "error": "Error de archivo"},
    "Deutsch": {"title": "Dokumentenkonverter", "service": "Konvertierung wählen:", "up": "Datei hier hochladen", "conv": "Datei konvertieren", "down": "Fertige Datei herunterladen", "wait": "Konvertiere...", "success": "Erfolgreich konvertiert!", "error": "Dateifehler"},
    "Português": {"title": "Conversor de Documentos", "service": "Escolher conversão:", "up": "Carregar arquivo aqui", "conv": "Converter arquivo", "down": "Baixar arquivo pronto", "wait": "Convertendo...", "success": "Convertido com sucesso!", "error": "Erro no arquivo"},
    "Italiano": {"title": "Convertitore di Documenti", "service": "Scegli la conversione:", "up": "Carica file qui", "conv": "Converti file", "down": "Scarica file pronto", "wait": "Conversione...", "success": "Convertito con successo!", "error": "Errore file"},
    "Русский": {"title": "Конвертер Документов", "service": "Выберите конвертацию:", "up": "Загрузите файл сюда", "conv": "Конвертировать", "down": "Скачать готовый файл", "wait": "Конвертация...", "success": "Успешно конвертировано!", "error": "Ошибка файла"},
    "中文": {"title": "文档和表格转换器", "service": "选择转换:", "up": "在此上传文件", "conv": "转换文件", "down": "下载准备好的文件", "wait": "正在转换...", "success": "转换成功!", "error": "文件错误"},
    "日本語": {"title": "ドキュメントコンバータ", "service": "変換を選択:", "up": "ここにファイルをアップロード", "conv": "ファイルを変換", "down": "準備完了ファイルをダウンロード", "wait": "変換中...", "success": "正常に変換されました!", "error": "ファイルエラー"},
    "한국어": {"title": "문서 변환기", "service": "변환 선택:", "up": "여기에 파일 업로드", "conv": "파일 변환", "down": "준비된 파일 다운로드", "wait": "변환 중...", "success": "성공적으로 변환되었습니다!", "error": "파일 오류"},
    "Türkçe": {"title": "Belge Dönüştürücü", "service": "Dönüşüm Seçin:", "up": "Dosyayı buraya yükle", "conv": "Dosyayı Dönüştür", "down": "Hazır Dosyayı İndir", "wait": "Dönüştürülüyor...", "success": "Başarıyla Dönüştürüldü!", "error": "Dosya Hatası"},
    "हिन्दी": {"title": "दस्तावेज़ कनवर्टर", "service": "रूपांतरण चुनें:", "up": "फ़ाइल यहाँ अपलोड करें", "conv": "फ़ाइल कनवर्ट करें", "down": "तैयार फ़ाइल डाउनलोड करें", "wait": "कनवर्ट हो रहा है...", "success": "सफलतापूर्वक कनवर्ट किया गया!", "error": "फ़ाइल त्रुटि"},
    "Nederlands": {"title": "Documenten Converter", "service": "Kies Conversie:", "up": "Upload bestand hier", "conv": "Converteer Bestand", "down": "Download Klaar Bestand", "wait": "Bezig met converteren...", "success": "Succesvol geconverteerd!", "error": "Bestandsfout"},
    "Polski": {"title": "Konwerter Dokumentów", "service": "Wybierz konwersję:", "up": "Prześlij plik tutaj", "conv": "Konwertuj plik", "down": "Pobierz gotowy plik", "wait": "Konwertowanie...", "success": "Pomyślnie przekonwertowano!", "error": "Błąd pliku"},
    "فارسی": {"title": "مبدل اسناد", "service": "تبدیل را انتخاب کنید:", "up": "فایل را اینجا آپلود کنید", "conv": "تبدیل فایل", "down": "دانلود فایل آماده", "wait": "در حال تبدیل...", "success": "با موفقیت تبدیل شد!", "error": "خطای فایل"},
    "Українська": {"title": "Конвертер Документів", "service": "Оберіть конвертацію:", "up": "Завантажте файл сюди", "conv": "Конвертувати", "down": "Завантажити готовий файл", "wait": "Конвертація...", "success": "Успішно конвертовано!", "error": "Помилка файлу"},
    "Indonesian": {"title": "Konverter Dokumen", "service": "Pilih Konversi:", "up": "Unggah File Di Sini", "conv": "Konversi File", "down": "Unduh File Siap", "wait": "Mengonversi...", "success": "Berhasil Dikonversi!", "error": "Kesalahan File"},
    "Tiếng Việt": {"title": "Chuyển đổi Tài liệu", "service": "Chọn Chuyển đổi:", "up": "Tải tệp lên tại đây", "conv": "Chuyển đổi Tệp", "down": "Tải Tệp Đã Xong", "wait": "Đang chuyển đổi...", "success": "Chuyển đổi Thành công!", "error": "Lỗi tệp"},
    "ไทย": {"title": "ตัวแปลงเอกสาร", "service": "เลือกการแปลง:", "up": "อัปโหลดไฟล์ที่นี่", "conv": "แปลงไฟล์", "down": "ดาวน์โหลดไฟล์ที่เสร็จแล้ว", "wait": "กำลังแปลง...", "success": "แปลงสำเร็จแล้ว!", "error": "ข้อผิดพลาดของไฟล์"}
}

# 3. إعدادات الألوان الصارمة (شريط أسود، خلفية صفراء)
st.markdown("""
    <style>
    /* الشاشة الرئيسية صفراء */
    .stApp { background-color: #fbc02d !important; }
    
    /* الشريط الجانبي أسود */
    [data-testid="stSidebar"] { background-color: #000000 !important; }
    
    /* نصوص الشريط الجانبي بيضاء */
    [data-testid="stSidebar"] * { color: white !important; font-weight: bold !important; }
    
    /* نصوص الشاشة الرئيسية سوداء لسهولة القراءة على الأصفر */
    h1, h2, h3, p, label, .stMarkdown { color: #000000 !important; font-weight: 900 !important; }
    
    /* تصميم الأزرار لتكون واضحة */
    div.stButton > button:first-child { background-color: #000000; color: #ffffff !important; font-size: 20px; border-radius: 10px; width: 100%; }
    div.stDownloadButton > button:first-child { background-color: #1b5e20; color: #ffffff !important; font-size: 20px; border-radius: 10px; width: 100%; }
    </style>
    """, unsafe_allow_html=True)

# 4. الشريط الجانبي (القائمة السوداء)
selected_lang = st.sidebar.selectbox("🌐", list(langs_dict.keys()))
t = langs_dict[selected_lang]

# القائمة الطويلة للتحويلات الفعلية المدعومة
conversion_options = [
    "PDF to Word",
    "PDF to Text",
    "Text to Word",
    "CSV to Excel",
    "Excel to CSV"
]
selected_service = st.sidebar.radio(t["service"], conversion_options)

# 5. الشاشة الرئيسية الصفراء (العنوان، الإحضار، التحويل، التنزيل)
st.title(t["title"])
st.subheader(f"🔄 {selected_service}")

# --- محرك التحويلات الفعلي (Backend Logic) ---

if selected_service == "PDF to Word":
    uploaded_file = st.file_uploader(t["up"], type=["pdf"])
    if uploaded_file:
        if st.button(t["conv"]):
            with st.spinner(t["wait"]):
                try:
                    # حفظ الملف مؤقتاً للتحويل
                    with open("temp.pdf", "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # التحويل الفعلي بواسطة مكتبة pdf2docx
                    cv = Converter("temp.pdf")
                    cv.convert("output.docx", start=0, end=None)
                    cv.close()
                    
                    st.success(t["success"])
                    with open("output.docx", "rb") as file_to_download:
                        st.download_button(t["down"], file_to_download, file_name="converted.docx")
                except Exception as e:
                    st.error(t["error"])

elif selected_service == "PDF to Text":
    uploaded_file = st.file_uploader(t["up"], type=["pdf"])
    if uploaded_file:
        if st.button(t["conv"]):
            with st.spinner(t["wait"]):
                try:
                    reader = PyPDF2.PdfReader(uploaded_file)
                    text = ""
                    for page in reader.pages:
                        text += page.extract_text() + "\n"
                    
                    st.success(t["success"])
                    st.download_button(t["down"], text, file_name="extracted.txt")
                except Exception as e:
                    st.error(t["error"])

elif selected_service == "Text to Word":
    uploaded_file = st.file_uploader(t["up"], type=["txt"])
    if uploaded_file:
        if st.button(t["conv"]):
            with st.spinner(t["wait"]):
                try:
                    content = uploaded_file.read().decode('utf-8')
                    doc = Document()
                    doc.add_paragraph(content)
                    
                    # حفظ في الذاكرة لتنزيله
                    bio = io.BytesIO()
                    doc.save(bio)
                    
                    st.success(t["success"])
                    st.download_button(t["down"], bio.getvalue(), file_name="document.docx")
                except Exception as e:
                    st.error(t["error"])

elif selected_service == "CSV to Excel":
    uploaded_file = st.file_uploader(t["up"], type=["csv"])
    if uploaded_file:
        if st.button(t["conv"]):
            with st.spinner(t["wait"]):
                try:
                    df = pd.read_csv(uploaded_file)
                    bio = io.BytesIO()
                    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    
                    st.success(t["success"])
                    st.download_button(t["down"], bio.getvalue(), file_name="data.xlsx")
                except Exception as e:
                    st.error(t["error"])

elif selected_service == "Excel to CSV":
    uploaded_file = st.file_uploader(t["up"], type=["xlsx"])
    if uploaded_file:
        if st.button(t["conv"]):
            with st.spinner(t["wait"]):
                try:
                    df = pd.read_excel(uploaded_file)
                    csv_data = df.to_csv(index=False).encode('utf-8')
                    
                    st.success(t["success"])
                    st.download_button(t["down"], csv_data, file_name="data.csv")
                except Exception as e:
                    st.error(t["error"])