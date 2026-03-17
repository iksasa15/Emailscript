#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""خادم ويب لواجهة إرسال الإيميلات"""

import os
import tempfile
import time
from flask import Flask, request, jsonify, send_from_directory

from send_emails import load_emails_from_excel, send_email, SMTP_SERVER, SMTP_PORT

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/send", methods=["POST"])
def send():
    sender_email = request.form.get("email", "").strip()
    sender_password = request.form.get("password", "")
    subject = request.form.get("subject", "").strip()
    body = request.form.get("body", "").strip()
    email_column = request.form.get("email_column", "email").strip() or "email"
    delay = float(request.form.get("delay", "1.5") or "1.5")

    if not sender_email or not sender_password:
        return jsonify({"ok": False, "error": "الإيميل وكلمة المرور مطلوبان"}), 400
    if not subject:
        return jsonify({"ok": False, "error": "عنوان الرسالة مطلوب"}), 400

    file = request.files.get("emails_file")
    if not file or file.filename == "":
        return jsonify({"ok": False, "error": "يرجى رفع ملف الإيميلات (إكسل)"}), 400
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"ok": False, "error": "الملف يجب أن يكون إكسل (.xlsx)"}), 400

    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name
        try:
            emails = load_emails_from_excel(tmp_path, email_column)
        finally:
            os.unlink(tmp_path)

        if not emails:
            return jsonify({"ok": False, "error": "لم يتم العثور على إيميلات صالحة في الملف"}), 400

        # تجهيز المرفقات (اختياري)
        attachments = []
        for f in request.files.getlist("attachments"):
            if f and f.filename:
                attachments.append((f.filename, f.read()))

        success = 0
        failed = 0
        errors = []

        for i, to_email in enumerate(emails):
            try:
                send_email(
                    to_email,
                    subject,
                    body,
                    sender_email,
                    sender_password,
                    attachments=attachments if attachments else None,
                )
                success += 1
            except Exception as e:
                failed += 1
                errors.append({"email": to_email, "message": str(e)})
            if i < len(emails) - 1 and delay > 0:
                time.sleep(delay)

        return jsonify({
            "ok": True,
            "total": len(emails),
            "success": success,
            "failed": failed,
            "errors": errors[:20],
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
