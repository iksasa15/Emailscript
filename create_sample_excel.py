#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""إنشاء ملف إكسل فاضي (قالب) لوضع الإيميلات فيه"""

from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Emails"

# الصف الأول فقط = عنوان العمود (مطلوب حتى يتعرّف السكربت على عمود الإيميل)
ws["A1"] = "email"
# من الصف الثاني فما تحت ضع عناوين الإيميل (واحد في كل صف)

wb.save("emails.xlsx")
print("تم إنشاء ملف emails.xlsx (فاضي). افتحه وضع الإيميلات في العمود A من الصف 2 وما تحت.")
