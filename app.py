from flask import Flask, render_template, request, send_file
import pdfkit
import io
from datetime import datetime
import os
import sys
import webbrowser
from threading import Timer

app = Flask(__name__)

# # wkhtmltopdf 路径
# wkhtmltopdf_path = r"D:\html-pdf\wkhtmltopdf\bin\wkhtmltopdf.exe"
# config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS  # 打包后的临时目录
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

wkhtmltopdf_path = os.path.join(BASE_DIR, "wkhtmltopdf", "bin", "wkhtmltopdf.exe")
config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)




# 公司信息
COMPANY = {
    "name": "ANYEE LIVING PTY LTD",
    "address": "57 Beecroft Cres, Templestowe VIC 3106",
    "email": "info@anyeeliving.com.au",
    "phone": "+61 466 019 988",
    "abn": "78688279094",
    "logo_path": "file:///" + os.path.abspath("static/logo.png").replace("\\", "/"),

    "currency": "AUD",
}


@app.route("/", methods=["GET"])
def form():
    today = datetime.today().strftime("%d %b, %Y")
    default_invoice_no = datetime.now().strftime("ANYEE#%Y%m%d001")
    return render_template("form.html", today=today, default_invoice_no=default_invoice_no)

@app.route("/generate", methods=["POST"])
def generate():
    form_data = request.form

    # 1. item的明细行
    items = []
    index = 0
    while True:
        desc_key = f"items[{index}][description]"
        if desc_key not in form_data or not form_data.get(desc_key, "").strip():
            break

        qty = float(form_data.get(f"items[{index}][quantity]", 0) or 0)
        unit_price = float(form_data.get(f"items[{index}][unit_price]", 0) or 0)
        tax_rate = float(form_data.get(f"items[{index}][tax_rate]", 0) or 0)

        # 原价 & 折扣说明（只用于显示，不参与计算）
        original_price_raw = form_data.get(f"items[{index}][original_price]", "").strip()
        original_price = float(original_price_raw) if original_price_raw else None
        discount_note = form_data.get(f"items[{index}][discount_note]", "").strip() or None

        line_subtotal = qty * unit_price
        line_tax = line_subtotal * tax_rate
        line_total = line_subtotal + line_tax

        items.append({
            "description": form_data.get(desc_key),
            "colour": form_data.get(f"items[{index}][colour]", ""),
            "size": form_data.get(f"items[{index}][size]", ""),
            "materials": form_data.get(f"items[{index}][materials]", ""),
            "quantity": qty,
            "unit_price": unit_price,
            "tax_rate": tax_rate,
            "line_subtotal": line_subtotal,
            "line_tax": line_tax,
            "line_total": line_total,
            "original_price": original_price,
            "discount_note": discount_note,
        })
        index += 1

    # 汇总金额
    subtotal = sum(i["line_subtotal"] for i in items)
    tax_total = sum(i["line_tax"] for i in items)

    promotion = float(form_data.get("promotion", 0) or 0)
    shipping_fee = float(form_data.get("shipping_fee", 0) or 0)
    voucher = float(form_data.get("voucher", 0) or 0)

    total = subtotal + tax_total - promotion + shipping_fee
    paid_by_customer = total - voucher

    invoice_data = {
        "invoice_number": form_data.get("invoice_number") or datetime.now().strftime("ANYEE#%Y%m%d001"),
        "issue_date": form_data.get("issue_date"),
        "bill_to_name": form_data.get("bill_to_name"),
        "bill_to_address": form_data.get("bill_to_address"),
        "bill_to_phone": form_data.get("bill_to_phone"),
        "payment_method": form_data.get("payment_method"),
        "promotion": promotion,
        "shipping_fee": shipping_fee,
        "subtotal": subtotal,
        "tax_total": tax_total,
        "voucher": voucher,
        "total": total,
        "paid_by_customer": paid_by_customer,
    }

    # 渲染 HTML
    html_str = render_template(
        "invoice.html",
        company=COMPANY,
        invoice=invoice_data,
        items=items,
    )

    # HTML 转 PDF
    pdf_bytes = pdfkit.from_string(
    html_str,
    False,
    configuration=config,
    options={
        "encoding": "UTF-8",
        "page-size": "A4",
        "margin-top": "10mm",
        "margin-bottom": "10mm",
        "margin-left": "10mm",
        "margin-right": "10mm",
        "enable-local-file-access": "",
    },
    )


    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"{invoice_data['invoice_number']}.pdf",
    )

def open_browser():
    webbrowser.open("http://127.0.0.1:5000/")

if __name__ == "__main__":
    # 1 秒后自动打开浏览器
    Timer(1, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
