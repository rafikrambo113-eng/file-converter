# 🔄 محوّل الملفات — File Converter

موقع Streamlit بيحوّل بين صيغ المستندات والصور المختلفة.

## المميزات

**المستندات**
- Word (.docx) ⇄ PDF
- Excel (.xlsx) → PDF
- استخراج الجداول من PDF إلى Excel
- PowerPoint (.pptx) → PDF
- PDF → صور (كل صفحة صورة منفصلة)
- صور → PDF (دمج أكتر من صورة في ملف واحد)

**الصور**
- تحويل بين PNG / JPEG / WEBP / BMP / GIF / TIFF / ICO / PDF
- تغيير الحجم (اختياري)
- التحكم في الجودة عند التصدير لـ JPEG أو WEBP

## التشغيل محليًا

```bash
pip install -r requirements.txt
streamlit run app.py
```

> لازم LibreOffice يكون متثبت على الجهاز عشان تحويلات Word/Excel/PowerPoint ← PDF تشتغل.
> على أوبونتو: `sudo apt-get install libreoffice`

## النشر على Streamlit Cloud

1. ارفع الملفات (`app.py`, `requirements.txt`, `packages.txt`) على مستودع GitHub.
2. من [share.streamlit.io](https://share.streamlit.io) اختار "New app" واربطه بالمستودع.
3. ملف `packages.txt` هيخلي Streamlit Cloud يثبّت LibreOffice تلقائيًا (مطلوب لتحويلات Word/Excel/PowerPoint).
4. أول تشغيل ممكن ياخد وقت أطول شوية لحد ما LibreOffice يتثبت.

## ملاحظات

- تحويل PDF → Excel بيعتمد على استخراج جداول حقيقية (نصوص وخطوط) من الملف؛ لو الجدول عبارة عن صورة/سكانر، الاستخراج مش هيلاقي جداول.
- الملفات الكبيرة أو الصفحات الكتيرة ممكن تاخد وقت أطول في التحويل.
