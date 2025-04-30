from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import pandas as pd
from datetime import datetime

def truncate_text(text, max_chars):
    """Truncate text and add ellipsis if it exceeds max_chars."""
    text = text.strip()
    if len(text) > max_chars:
        return text[:max_chars].rstrip() + "..."
    return text

def generate_pdf(data, filename):
    if data['InsuranceID'] == '' or data['InsuranceID'] is None or data['InsuranceID'] == 'nan':
        print("No data to generate PDF.")
        return

    left_margin = right_margin = 5  # 0.5 inch
    page_width = letter[0]
    usable_width = page_width - left_margin - right_margin

    doc = SimpleDocTemplate(filename, pagesize=letter,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=30,
        bottomMargin=30)
    styles = getSampleStyleSheet()
    flowables = []

    # Title
    flowables.append(Paragraph("<b>Patient Report</b>", styles["Title"]))
    flowables.append(Spacer(1, 12))
    formatted_bdate = pd.to_datetime(data['Birthday']).strftime('%m/%d/%Y')
    formattedfirst_visit = pd.to_datetime(data['FirstServiceDate']).strftime('%m/%d/%Y')
    formattedlast_visit = pd.to_datetime(data['LastServiceDate']).strftime('%m/%d/%Y')
    # Patient Info
    patient_info = f"""
    <b>Name:</b> {data['FirstName']} {data['LastName']}<br/>
    <b>DOB:</b> {formatted_bdate}<br/>
    <b>Address:</b> {data['Address']}<br/>
    <b>Insurance:</b> {data['InsuranceProvider'].replace('z ','')}<br/>
    <b>First Visit:</b> {formattedfirst_visit}<br/>
    <b>Last Visit:</b> {formattedlast_visit}<br/>
    <b>Insurance Claim ID:</b> {data['InsuranceID']}<br/>
    <b>Medical Office:</b> {data['Facility']}<br/>
    """
    flowables.append(Paragraph(patient_info, styles["Normal"]))
    flowables.append(Spacer(1, 12))

    total_balance = f"""
    <b>Total Balance:</b> ${data['TotalBalance']:.2f}<br/>
    """

    # Table Header
    table_data = [
        ["Service Date", "Code", "Units", "Provider", "Balance"]
    ]

    # Table rows
    for case in data["ChargeDetails"]:
        # desc = truncate_text(case.get("description", ""), max_chars=8)
        #case_number = truncate_text(case.get("case_number", ""), max_chars=10)
        provider = truncate_text(case.get("provider", ""), max_chars=19)
        charge_date_str = pd.to_datetime(case['Visit']).strftime('%m/%d/%Y') if pd.notna(case['Visit']) else 'N/A'
        linebalancerounded = round(float(case.get("LineBalance", 0)), 2)
        row = [
            #case_number,
            charge_date_str,
            case.get("CPT", ""),
            case.get("Units", ""),
            #desc,
            case.get("Provider Name"),
            "$"+str(linebalancerounded) 
        ]
        table_data.append(row)

    # Table styling
    # Define original column width intentions
    original_col_widths = [60, 30, 20, 90, 80] # Adjusted description width and added location width
    total_width = sum(original_col_widths)

    # Scale down to fit the usable page width
    scale_factor = usable_width / total_width
    scaled_col_widths = [width * scale_factor for width in original_col_widths]

    table = Table(table_data, colWidths=scaled_col_widths, repeatRows=1) # Apply the scaled widths
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    flowables.append(table)
    flowables.append(Spacer(1, 12))
    flowables.append(Paragraph(total_balance.replace('[', '').replace(']','').replace("'",''), styles["Normal"]))
    doc.build(flowables)
    print(f"PDF generated: {filename}")


# <b>Lifetime Charges:</b> {data['Lifetime Charges']}<br/>
# <b>Lifetime Payments:</b> {data['Lifetime Payments']}<br/>
# <b>Lifetime Adjustments:</b> {data['Lifetime Adjustments']}<br/>
# <b>Lifetime Write Off:</b> {data['Lifetime Write Off']}<br/>
# <b>Lifetime Sequestration:</b> {data['Lifetime Sequestration']}<br/>