import streamlit as str_web
import speedtest
import socket
import psutil
import pandas as pd
import webbrowser
from datetime import datetime

# إعدادات الصفحة الافتراضية للموقع
str_web.set_page_config(page_title="حاسب تفاصيل الإنترنت الشامل", page_icon="📊", layout="centered")

# تأكيد فتح الموقع على متصفح Chrome تلقائياً عند التشغيل
try:
    chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe %s'
    browser = webbrowser.get(chrome_path)
except:
    browser = webbrowser

if "opened" not in str_web.session_state:
    browser.open("http://localhost:8501")
    str_web.session_state["opened"] = True

# 1. قاموس اللغات العشرة الكامل والصافي 100% للموقع
LANGUAGES = {
    "العربية": {
        "title": "📊 حاسب تفاصيل الإنترنت الشامل",
        "type": "نوع الاتصال: ", "usage": "الاستهلاك الحالي: ", "download_label": "تحميل", "upload_label": "رفع",
        "ip": "عنوان الـ IP المحلي: ", "ping": "زمن الاستجابة (Ping): ", "down_speed": "سرعة التحميل (Download): ",
        "up_speed": "سرعة الرفع (Upload): ", "status_idle": "اضغط على الزر لبدء الفحص وحساب الاستهلاك",
        "status_scanning": "جاري الفحص... يرجى الانتظار قليلاً", "status_success": "تم الفحص بنجاح!",
        "btn_scan": "ابدأ الفحص الآن", "btn_download": "📥 تنزيل جدول البيانات (Excel)",
        "wifi": "واي فاي (Wi-Fi)", "ethernet": "كابل (Ethernet)", "unknown": "غير معروف"
    },
    "English": {
        "title": "📊 Comprehensive Internet Calculator",
        "type": "Connection Type: ", "usage": "Current Usage: ", "download_label": "Download", "upload_label": "Upload",
        "ip": "Local IP Address: ", "ping": "Ping (Latency): ", "down_speed": "Download Speed: ",
        "up_speed": "Upload Speed: ", "status_idle": "Click button to start scan",
        "status_scanning": "Scanning... Please wait a moment", "status_success": "Scan completed successfully!",
        "btn_scan": "Start Scan Now", "btn_download": "📥 Download Data Table (Excel)",
        "wifi": "Wi-Fi", "ethernet": "Ethernet", "unknown": "Unknown"
    },
    "Mandarin (الصينية)": {
        "title": "📊 综合 互联网 计算器",
        "type": "连接类型: ", "usage": "当前流量: ", "download_label": "下载", "upload_label": "上传",
        "ip": "本地 IP 地址: ", "ping": "延迟 (Ping): ", "down_speed": "下载速度: ",
        "up_speed": "上传速度: ", "status_idle": "点击按钮开始扫描",
        "status_scanning": "正在扫描... 请稍候", "status_success": "扫描成功完成！",
        "btn_scan": "立即开始扫描", "btn_download": "📥 下载数据表 (Excel)",
        "wifi": "无线网络 (Wi-Fi)", "ethernet": "以太网 (Ethernet)", "unknown": "未知"
    },
    "Hindi (الهندية)": {
        "title": "📊 व्यापक इंटरनेट कैलकुलेटर",
        "type": "कनेक्शन का प्रकार: ", "usage": "वर्तमान उपयोग: ", "download_label": "डाउनलोड", "upload_label": "अपलोड",
        "ip": "स्थानीय आईपी पता: ", "ping": "पिंग (Ping): ", "down_speed": "डाउनलोड गति: ",
        "up_speed": "अपलोड गति: ", "status_idle": "स्कैन शुरू करने के لیے بٹن دبائیں",
        "status_scanning": "स्कैनिंग जारी है... कृपया प्रतीक्षा करें", "status_success": "स्कैन सफलतापूर्वक पूरा हुआ!",
        "btn_scan": "اب اسکین شروع کریں", "btn_download": "📥 ڈیٹا نامہ ڈاؤن لوڈ کریں (Excel)",
        "wifi": "वाई-फाई", "ethernet": "ईथरनेट", "unknown": "अज्ञात"
    },
    "Español (الإسبانية)": {
        "title": "📊 Calculadora Integral de Internet",
        "type": "Tipo de conexión: ", "usage": "Consumo actual: ", "download_label": "Descarga", "upload_label": "Subida",
        "ip": "Dirección IP local: ", "ping": "Latencia (Ping): ", "down_speed": "Velocidad de descarga: ",
        "up_speed": "Velocidad de subida: ", "status_idle": "Presione el botón para iniciar el escaneo",
        "status_scanning": "Escaneando... Por favor espere un momento", "status_success": "¡Escaneo completado con éxito!",
        "btn_scan": "Iniciar escaneo ahora", "btn_download": "📥 Descargar Tabla de Datos (Excel)",
        "wifi": "Wi-Fi", "ethernet": "Ethernet", "unknown": "Desconocido"
    },
    "Français (الفرنسية)": {
        "title": "📊 Calculateur Internet Complet",
        "type": "Type de connexion: ", "usage": "Consommation actuelle: ", "download_label": "Téléchargement", "upload_label": "Téléversement",
        "ip": "Adresse IP locale: ", "ping": "Latence (Ping): ", "down_speed": "Vitesse de téléchargement: ",
        "up_speed": "Vitesse de téléversement: ", "status_idle": "Cliquez sur le bouton pour lancer le scan",
        "status_scanning": "Scan en cours... Veuillez patienter un instant", "status_success": "Scan réussi !",
        "btn_scan": "Lancer le scan maintenant", "btn_download": "📥 Télécharger le Tableau (Excel)",
        "wifi": "Wi-Fi", "ethernet": "Ethernet", "unknown": "Inconnu"
    },
    "Bengali (البنغالية)": {
        "title": "📊 ব্যাপক ইন্টারনেট ক্যালকুলেটর",
        "type": "সংযোগের ধরন: ", "usage": "বর্তমান ব্যবহার: ", "download_label": "ডাউনলোড", "upload_label": "আপলোড",
        "ip": "স্থানীয় আইپی ঠিকানা: ", "ping": "পিং (Ping): ", "down_speed": "ডাউনলোড গতি: ",
        "up_speed": "আপロード গতি: ", "status_idle": "স্ক্যান শুরু করতে বোতাম টিপুন",
        "status_scanning": "স্ক্যান হচ্ছে... অনুগ্রহ করে একটু অপেক্ষা করুন", "status_success": "স্ক্যান সফলভাবে সম্পন্ন হয়েছে!",
        "btn_scan": "এখনই স্ক্যান শুরু করুন", "btn_download": "📥 ডাটা টেবিল ডাউনলোড করুন (Excel)",
        "wifi": "ওয়াই-ফাই", "ethernet": "ইথারনেট", "unknown": "অজানা"
    },
    "Português (البرتغالية)": {
        "title": "📊 Calculadora de Internet Abrangente",
        "type": "Tipo de conexão: ", "usage": "Consumo actual: ", "download_label": "Download", "upload_label": "Upload",
        "ip": "Endereço IP local: ", "ping": "Latência (Ping): ", "down_speed": "Velocidade de download: ",
        "up_speed": "Velocidade de upload: ", "status_idle": "Clique no botão para iniciar a verificação",
        "status_scanning": "Verificando... Por favor, aguarde um momento", "status_success": "Verificação concluída com sucesso!",
        "btn_scan": "Iniciar verificação agora", "btn_download": "📥 Baixar Tabela de Datos (Excel)",
        "wifi": "Wi-Fi", "ethernet": "Ethernet", "unknown": "Desconhecido"
    },
    "Russian (الروسية)": {
        "title": "📊 Комплексный Интернет-Калькулятор",
        "type": "Тип подключения: ", "usage": "Текущий расход: ", "download_label": "Загрузка", "upload_label": "Отдача",
        "ip": "Локальный IP-адрес: ", "ping": "Задержка (Ping): ", "down_speed": "Скорость загрузки: ",
        "up_speed": "Скорость отдачи: ", "status_idle": "Нажмите кнопку для начала проверки",
        "status_scanning": "Идет проверка... Пожалуйста, подождите немного", "status_success": "Проверка успешно завершена!",
        "btn_scan": "Начать проверку", "btn_download": "📥 Скачать таблицу данных (Excel)",
        "wifi": "Wi-Fi", "ethernet": "Ethernet", "unknown": "Неизвестно"
    },
    "Urdu (الأردية)": {
        "title": "📊 جامع انٹرنیٹ کیلکولیٹر",
        "type": "کنکشن کی قسم: ", "usage": "موجودہ استعمال: ", "download_label": "ڈاؤن لوڈ", "upload_label": "اپ لوڈ",
        "ip": "مقامی IP ایڈريس: ", "ping": "پنگ (Ping): ", "down_speed": "ڈاؤن لوڈ کی رفتار: ",
        "up_speed": "اپ لوڈ کی رفتار: ", "status_idle": "اسکین شروع کرنے کے لیے بٹن دبائیں",
        "status_scanning": "اسکیننگ جاری ہے... براہ کرم تھوڑا انتظار کریں", "status_success": "اسکین کامیابی سے مکمل ہو گیا!",
        "btn_scan": "ابھی اسکین شروع کریں", "btn_download": "📥 ڈیٹا ٹیبل ڈاؤن لوڈ کریں (Excel)",
        "wifi": "وائی فائی", "ethernet": "ایتھرنیٹ", "unknown": "نامعلوم"
    }
}

# 2. قوائم الألوان العشرة الصافية لتتبع اللغات
THEMES = {
    "العربية": ["الأسود", "الأحمر", "الأزرق", "الأصفر", "الأخضر", "البرتقالي", "الرمادي", "البنفسجي", "البني", "الأبيض"],
    "English": ["Black", "Red", "Blue", "Yellow", "Green", "Orange", "Gray", "Purple", "Brown", "White"],
    "Mandarin (الصينية)": ["黑色", "红色", "蓝色", "黄色", "绿色", "橙色", "灰色", "紫色", "棕色", "白色"],
    "Hindi (الهندية)": ["काला", "लाल", "नीला", "पीला", "हरा", "नारंगी", "धूसर", "बैंगनी", "भूरा", "सफेद"],
    "Español (الإسبانية)": ["Negro", "Rojo", "Azul", "Amarillo", "Verde", "Naranja", "Gris", "Púrpura", "Marrón", "Blanco"],
    "Français (الفرنسية)": ["Noir", "Rouge", "Bleu", "Jaune", "Vert", "Orange", "Gris", "Violet", "Marron", "Blanc"],
    "Bengali (البنغالية)": ["কালো", "লাল", "নীল", "হলুদ", "সবুজ", "কমলা", "ধূসর", "বেগুনী", "বাদামী", "সাদা"],
    "Português (البرتغالية)": ["Preto", "Vermelho", "Azul", "Amarelo", "Verde", "Laranja", "Cinza", "Roxo", "Marrom", "Branco"],
    "Russian (الروسية)": ["Черный", "Красный", "Синий", "Желтый", "Зеленый", "Оранжевый", "Серый", "Пурпурный", "Коричневый", "Белый"],
    "Urdu (الأردية)": ["کالا", "سرخ", "نیلا", "پیلا", "ہرا", "نارنجی", "گرے", "جامنی", "بھورا", "سفید"]
}

COLOR_CODES = {
    "الأسود": "#121212", "Black": "#121212", "黑色": "#121212", "काला": "#121212", "Negro": "#121212", "Noir": "#121212", "কালো": "#121212", "Preto": "#121212", "Черный": "#121212", "کالا": "#121212",
    "الأحمر": "#d32f2f", "Red": "#d32f2f", "红色": "#d32f2f", "लाल": "#d32f2f", "Rojo": "#d32f2f", "Rouge": "#d32f2f", "লাল": "#d32f2f", "Vermelho": "#d32f2f", "Красный": "#d32f2f", "سرخ": "#d32f2f",
    "الأزرق": "#1976d2", "Blue": "#1976d2", "蓝色": "#1976d2", "नीला": "#1976d2", "Azul": "#1976d2", "Bleu": "#1976d2", "নীল": "#1976d2", "Синий": "#1976d2", "نیلا": "#1976d2",
    "الأصفر": "#fbc02d", "Yellow": "#fbc02d", "黄色": "#fbc02d", "पीला": "#fbc02d", "Amarillo": "#fbc02d", "Jaune": "#fbc02d", "হলুদ": "#fbc02d", "Amarelo": "#fbc02d", "Желтый": "#fbc02d", "پیلا": "#fbc02d",
    "الأخضر": "#388e3c", "Green": "#388e3c", "绿色": "#388e3c", "हरा": "#388e3c", "Verde": "#388e3c", "Vert": "#388e3c", "সবুজ": "#388e3c", "Зеленый": "#388e3c", "ہرا": "#388e3c",
    "البرتقالي": "#f57c00", "Orange": "#f57c00", "橙色": "#f57c00", "नारंगी": "#f57c00", "Naranja": "#f57c00", "কমলা": "#f57c00", "Laranja": "#f57c00", "Оранжевый": "#f57c00", "نارنجی": "#f57c00",
    "الرمادي": "#616161", "Gray": "#616161", "灰色": "#616161", "धूसر": "#616161", "Gris": "#616161", "ধূসর": "#616161", "Cinza": "#616161", "Серый": "#616161", "گرے": "#616161",
    "البنفسجي": "#7b1fa2", "Purple": "#7b1fa2", "紫色": "#7b1fa2", "बैंगنی": "#7b1fa2", "Púrpura": "#7b1fa2", "Violet": "#7b1fa2", "বেগুনী": "#7b1fa2", "Roxo": "#7b1fa2", "Пурпурный": "#7b1fa2", "جامنی": "#7b1fa2",
    "البني": "#5d4037", "Brown": "#5d4037", "棕色": "#5d4037", "भूра": "#5d4037", "Marrón": "#5d4037", "Marron": "#5d4037", "বাদামী": "#5d4037", "Marrom": "#5d4037", "Коричневый": "#5d4037", "بھورا": "#5d4037",
    "الأبيض": "#f5f5f5", "White": "#f5f5f5", "白色": "#f5f5f5", "सफेद": "#f5f5f5", "Blanco": "#f5f5f5", "Blanc": "#f5f5f5", "সাদা": "#f5f5f5", "Branco": "#f5f5f5", "Белый": "#f5f5f5", "سفید": "#f5f5f5"
}

col1, col2 = str_web.columns(2)
with col1:
    lang_choice = str_web.selectbox("Language / اللغة", list(LANGUAGES.keys()))
with col2:
    color_choice = str_web.selectbox("Theme Color / اللون", THEMES[lang_choice])

trans = LANGUAGES[lang_choice]
selected_color = COLOR_CODES[color_choice]

# تحديد لون النص التوضيحي بذكاء حسب الخلفية (أسود للثيمات الفاتحة، وأبيض للثيمات الغامقة)
text_box_bg = "rgba(0,0,0,0.6)" if color_choice in ["الأصفر", "الأبيض", "Yellow", "White", "白色", "सफेद", "Amarillo", "Blanco", "Jaune", "Blanc", "হলুদ", "সাদা", "Amarelo", "Желтый", "Белый", "پیلا", "سفید"] else "rgba(255,255,255,0.15)"
text_color = "black" if color_choice in ["الأصفر", "الأبيض", "Yellow", "White"] else "white"

# هندسة وتعديل تصميم الـ CSS لتحسين حجم ولون النص وتنسيق ألوان الأزرار
str_web.markdown(f"""
    <style>
    .stApp {{ background-color: {selected_color}; color: white; }}
    div[data-testid="stMetricValue"] {{ color: white !important; }}
    
    /* تعديل تصميم النص التوضيحي ليكون عريضاً، كبيراً وواضحاً جداً */
    .custom-status-box {{
        background-color: {text_box_bg};
        color: white !important;
        padding: 15px;
        border-radius: 8px;
        font-size: 20px !important;
        font-weight: bold !important;
        text-align: center;
        margin-top: 15px;
        border: 1px solid rgba(255,255,255,0.2);
    }}
    
    /* تعديل تصميم أزرار الويب لتظهر النصوص بوضوح داخلها */
    .stButton>button {{
        width: 100% !important;
        background-color: white !important;
        color: black !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 8px !important;
        padding: 10px !important;
    }}
    </style>
    """, unsafe_allow_html=True)

str_web.title(trans["title"])
str_web.write("---")

def get_network_info():
    conn_type = "unknown"
    for interface, info in psutil.net_if_stats().items():
        if info.isup and "loopback" not in interface.lower():
            if "wi-fi" in interface.lower() or "wlan" in interface.lower():
                conn_type = "wifi"
                break
            elif "ethernet" in interface.lower() or "ether" in interface.lower():
                conn_type = "ethernet"
                break
    io = psutil.net_io_counters()
    return conn_type, io.bytes_recv / (1024**3), io.bytes_sent / (1024**3)

# زر الفحص الرئيسي (الآن النص يظهر بداخله بوضوح)
if str_web.button(trans["btn_scan"]):
    with str_web.spinner(trans["status_scanning"]):
        try:
            hostname = socket.gethostname()
            local_ip = socket.gethostbyname(hostname)
            conn_type, down_gb, up_gb = get_network_info()
            
            st = speedtest.Speedtest()
            st.get_best_server()
            
            download_speed = f"{st.download() / 1_000_000:.2f} Mbps"
            upload_speed = f"{st.upload() / 1_000_000:.2f} Mbps"
            ping = f"{st.results.ping:.1f} ms"
            type_str = trans[conn_type]
            
            str_web.success(trans["status_success"])
            
            str_web.metric(label=trans["type"], value=type_str)
            str_web.metric(label=trans["usage"], value=f"{trans['download_label']} {down_gb:.2f} GB | {trans['upload_label']} {up_gb:.2f} GB")
            str_web.metric(label=trans["ip"], value=local_ip)
            str_web.metric(label=trans["ping"], value=ping)
            str_web.metric(label=trans["down_speed"], value=download_speed)
            str_web.metric(label=trans["up_speed"], value=upload_speed)
            
            data_to_export = {
                "Metric / القياس": [trans["type"], trans["usage"], trans["ip"], trans["ping"], trans["down_speed"], trans["up_speed"], "وقت الفحص"],
                "Value / القيمة": [type_str, f"{down_gb:.2f} GB / {up_gb:.2f} GB", local_ip, ping, download_speed, upload_speed, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
            }
            df = pd.DataFrame(data_to_export)
            
            df.to_excel("stats.xlsx", index=False)
            with open("stats.xlsx", "rb") as f:
                str_web.download_button(
                    label=trans["btn_download"],
                    data=f,
                    file_name=f"Internet_Stats_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            str_web.error(f"Error: {e}")
else:
    # عرض العبارة التوضيحية داخل صندوق مخصص ومحسن ليظهر الخط كبيراً وعريضاً
    str_web.markdown(f'<div class="custom-status-box">{trans["status_idle"]}</div>', unsafe_allow_html=True)