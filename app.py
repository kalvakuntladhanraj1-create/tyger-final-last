from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
from datetime import datetime

app = Flask(__name__)


# ---------------- HOME ----------------
@app.route("/")
def home():
    return render_template("sale.html")


# ---------------- GENERATE REPORT ----------------
@app.route("/generate_sale", methods=["POST"])
def generate_sale():

    # -------- BASIC --------
    context = {
        "DATE": request.form.get("DATE"),
        "BRANCH_NAME": request.form.get("BRANCH_NAME"),
        "BRANCH_AREA": request.form.get("BRANCH_AREA"),
        "APPLICANT_BOLD_CAPS": request.form.get("APPLICANT_BOLD_CAPS"),
        "APPLICATION_NO": request.form.get("APPLICATION_NO"),

        # -------- PROPERTY --------
        "DOOR_NO": request.form.get("DOOR_NO"),
        "PLOT_NO": request.form.get("PLOT_NO"),
        "ASSESSMENT_NO": request.form.get("ASSESSMENT_NO"),
        "EXTENT_YARDS": request.form.get("EXTENT_YARDS"),
        "SURVEY_NO": request.form.get("SURVEY_NO"),
        "ADDRESS": request.form.get("ADDRESS"),
        "VILLAGE": request.form.get("VILLAGE"),
        "GRAM_PANCHAYAT": request.form.get("GRAM_PANCHAYAT"),
        "MANDAL": request.form.get("MANDAL"),
        "SRO": request.form.get("SRO"),
        "RO": request.form.get("RO"),
        "DISTRICT": request.form.get("DISTRICT"),

        # -------- EXTRA --------
        "POSSESSION_DATE": request.form.get("POSSESSION_DATE"),
        "FROM_YEARS": request.form.get("FROM_YEARS"),
        "POSSESSION_NAME": request.form.get("POSSESSION_NAME"),

        "H_T_DATE": request.form.get("H_T_DATE"),
        "HOUSE_TAX_RECIPT_NO": request.form.get("HOUSE_TAX_RECIPT_NO"),
        "FINANCIAL_YEARS": request.form.get("FINANCIAL_YEARS"),
        "HOUSE_TAX_NAME": request.form.get("HOUSE_TAX_NAME"),
        "HOUSE_TAX_ISSUED_BY": request.form.get("HOUSE_TAX_ISSUED_BY"),

        "EC_DATE": request.form.get("EC_DATE"),
        "EC_NO": request.form.get("EC_NO"),
    }

    # -------- BOOLEAN --------
    context["HAS_ELECTRICITY_BILL"] = request.form.get("HAS_ELECTRICITY_BILL") == "true"
    context["HAS_MORTGAGE"] = request.form.get("HAS_MORTGAGE") == "true"

    # -------- ELECTRICITY --------
    context["ELECTRICITY_BILL_DATE"] = request.form.get("ELECTRICITY_BILL_DATE")
    context["SERVICE_NO"] = request.form.get("SERVICE_NO")
    context["ELECTRICITY_NAME"] = request.form.get("ELECTRICITY_NAME")

    # -------- MORTGAGE --------
    context["MORTGAGE_DEED_NO"] = request.form.get("MORTGAGE_DEED_NO")
    context["MORTGAGE_DEED_DATE"] = request.form.get("MORTGAGE_DEED_DATE")
    context["MORTGAGE_COMPANY"] = request.form.get("MORTGAGE_COMPANY")

    # -------- DOCUMENTS --------
    documents = []

    types = request.form.getlist("doc_type[]")
    numbers = request.form.getlist("doc_number[]")
    dates = request.form.getlist("doc_date[]")
    executants = request.form.getlist("doc_executant[]")
    owners = request.form.getlist("doc_owner[]")
    relations = request.form.getlist("doc_relation[]")
    worths = request.form.getlist("doc_worth[]")

    for i in range(len(types)):
        # Safe check (IMPORTANT)
        if types[i] and numbers[i] and dates[i]:
            documents.append({
                "type": types[i],
                "number": numbers[i],
                "date": dates[i],
                "executant": executants[i],
                "owner": owners[i],
                "relation": relations[i],
                "worth": worths[i]
            })

    context["DOCUMENTS"] = documents

    # -------- LOAD TEMPLATE --------
    doc = DocxTemplate("templates_docx/tyger_report.docx")
    doc.render(context)

    # -------- SAVE FILE --------
    filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = os.path.join("templates_docx", filename)
    doc.save(output_path)

    # -------- DOWNLOAD --------
    return send_file(output_path, as_attachment=True)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)