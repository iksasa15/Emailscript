# سكربت إرسال الإيميلات من إكسل

سكربت بايثون يقرأ قائمة إيميلات من ملف إكسل ويرسل لها رسالة واحدة.

## التثبيت

```bash
pip install -r requirements.txt
```

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

## التشغيل

```bash
python send_emails.py
```

يرسل السكربت رسالة واحدة لكل إيميل في العمود مع انتظار بضع ثوانٍ بين كل إرسال (يمكن تغييرها عبر `DELAY_BETWEEN_EMAILS`).
