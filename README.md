# سكربت إرسال الإيميلات من إكسل

سكربت بايثون يقرأ قائمة إيميلات من ملف إكسل ويرسل لها رسالة واحدة.

## التثبيت والتشغيل (الواجهة)

```bash
cd /Users/ahmed/Desktop/Emailscript
pip install -r requirements.txt
python app.py
```

ثم افتح المتصفح على: **http://127.0.0.1:5000**

للتطوير مع وضع التصحيح (اختياري):

```bash
FLASK_DEBUG=1 python app.py
```

## الاستضافة (Railway / Render / Fly.io / Heroku)

المشروع جاهز للتشغيل بـ **Gunicorn** عبر ملف `Procfile`.

1. ارفع المشروع إلى **GitHub** ثم اربطه بالمنصة.
2. **أمر البناء:** `pip install -r requirements.txt` (غالباً يُكتشف تلقائياً).
3. **أمر التشغيل:** يُقرأ من `Procfile` (`gunicorn ... app:app`).
4. المنصة تضبط متغير **`PORT`** تلقائياً؛ لا حاجة لتعديل الكود.
5. **`runtime.txt`** يحدد إصدار بايثون (يمكن تغييره حسب المنصة).

**مهم:** لا تفعّل `FLASK_DEBUG` على السيرفر. التطبيق يرسل إيميلات — استخدم HTTPS (المنصة توفره)، ولا تشارك رابط المشروع علناً إن لم تضف حماية (كلمة مرور / IP).

**بدائل:** [PythonAnywhere](https://www.pythonanywhere.com) (إعداد يدوي لـ WSGI)، أو **VPS** مع Nginx أمام Gunicorn:

```bash
gunicorn --bind 127.0.0.1:8000 --timeout 600 --workers 1 app:app
```

## النشر على Render

1. ارفع المشروع إلى **GitHub** (مستودع عام أو خاص).
2. سجّل الدخول إلى [dashboard.render.com](https://dashboard.render.com) واختر **New**:
   - **Blueprint** (موصى به): اربط الريبو واختر ملف `render.yaml` — يُنشئ خدمة ويب تلقائياً.
   - أو **Web Service**: اربط نفس الريبو وعيّن يدوياً:
     - **Build Command:** `pip install -r requirements.txt`
     - **Start Command:** `gunicorn --bind 0.0.0.0:$PORT --workers 1 --threads 2 --timeout 600 app:app`
3. إصدار بايثون: يُضبط عبر `PYTHON_VERSION` في `render.yaml` أو ملف `.python-version` في جذر المشروع ([توثيق Render](https://render.com/docs/python-version)).
4. بعد اكتمال النشر، افتح الرابط `https://اسم-الخدمة.onrender.com`.

**ملاحظات:** الخطة المجانية قد تُدخل الخدمة في **سبات** بعد فترة بدون زيارات؛ أول طلب بعد السبات قد يستغرق عشرات الثواني. الإرسال الطويل للإيميلات مدعوم بفضل `--timeout 600`.

## ملف الإكسل

- اسم الملف الافتراضي: `emails.xlsx`
- يجب أن يحتوي على عمود اسمه `email` (أو غيّر `EMAIL_COLUMN` في السكربت).
- الصف الأول = عناوين الأعمدة، من الصف الثاني تبدأ الإيميلات.

مثال:

| email              |
|--------------------|
| user1@example.com  |
| user2@example.com  |

لإنشاء ملف نموذجي:

```bash
python create_sample_excel.py
```

## الإعدادات

افتح `send_emails.py` وعدّل:

- `EXCEL_FILE` — مسار ملف الإكسل
- `EMAIL_COLUMN` — اسم عمود الإيميل أو رقمه (مثلاً 1 للعمود A)
- `SENDER_EMAIL` — إيميلك
- `EMAIL_SUBJECT` — موضوع الرسالة
- `EMAIL_BODY` — نص الرسالة

## كلمة المرور (Gmail)

1. فعّل "التحقق بخطوتين" من إعدادات حساب Google.
2. من [حساب Google → الأمان → كلمات مرور التطبيقات](https://myaccount.google.com/apppasswords) أنشئ "كلمة مرور تطبيق".
3. استخدمها في المتغير `EMAIL_PASSWORD` (ولا تضع كلمة مرورك العادية).

تشغيل مع كلمة المرور من المتغير:

```bash
export EMAIL_PASSWORD='كلمة_مرور_التطبيق'
python send_emails.py
```

أو ضعها مؤقتاً في السكربت في `SENDER_PASSWORD` (غير مستحسن).

## التشغيل من الطرفية (بدون واجهة)

```bash
cd /Users/ahmed/Desktop/Emailscript
python send_emails.py
```

يرسل السكربت رسالة واحدة لكل إيميل في العمود مع انتظار بضع ثوانٍ بين كل إرسال (يمكن تغييرها عبر `DELAY_BETWEEN_EMAILS`).
