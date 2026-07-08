import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from jinja2 import Environment, FileSystemLoader
from playwright.sync_api import sync_playwright
import sys
from pathlib import Path

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except Exception:
        base_path = Path(".")

    return base_path / relative_path

def text(value, default=" "):
    return value if pd.notna(value) else default

def money(value):
    if pd.isna(value):
        return "$0.00"

    try:
        value = float(str(value).replace("$", "").replace(",", ""))
        return f"${value:,.2f}"
    except (ValueError, TypeError):
        return "$0.00"

def generate_documents(df, template_name, output_name, env):
    output_dir = Path("generated pdfs")
    output_dir.mkdir(exist_ok=True)

    template = env.get_template(template_name)

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome") 

        for _, row in df.iterrows():

            html_content = template.render(
                shareholder_details=text(row["Shareholder Details"]),
                shareholder_name=text(row["Shareholder Name"]),
                shareholder_address=text(row["Shareholder Address"]),
                shareholder_email=text(row["Shareholder Email"]),

                spv_name=text(row["SPV Name"]),
                spv_reg_num=text(row["SPV Reg Num"]),
                spv_director_name=text(row["SPV Director Name"]),

                company_name=text(row["Company Name"]),
                company_reg_num=text(row["Company Reg Num"]),

                currency=text(row["Currency"]),
                conversion_desc=text(row["Conversion Description"], ""),

                date=row["Date"].strftime("%d/%m/%Y") if pd.notna(row["Date"]) else "-",
                due_date_format1=row["Due Date"].strftime("%A %d %B %Y") if pd.notna(row["Due Date"]) else "-",
                due_date_format2=row["Due Date"].strftime("%d/%m/%Y") if pd.notna(row["Due Date"]) else "-",

                one_time_fees=money(row["One Time Fees"]),
                admin_service_fees=money(row["Admin Service Fees"]),

                rate=(
                    money(float(str(row["Admin Service Fees"]).replace("$", "").replace(",", "")) / 5)
                    if pd.notna(row["Admin Service Fees"])
                    else "$0.00"
                ),

                one_time_fee_quantity=1 if pd.notna(row["One Time Fees"]) else "0",
                admin_service_fee_quantity=5 if pd.notna(row["Admin Service Fees"]) else "0",

                miscellaneous_desc=text(row["Miscellaneous Description"], ""),
                miscellaneous_fees=money(row["Miscellaneous Fees"]),

                total_fees=money(row["Total Fees"]),
                hard_commit_sum=money(row["Hard Commit Sum"]),

                total_amount = money(
                    (float(str(row["Hard Commit Sum"]).replace("$", "").replace(",", ""))
                    if pd.notna(row["Hard Commit Sum"]) else 0)
                    +
                    (float(str(row["Total Fees"]).replace("$", "").replace(",", ""))
                    if pd.notna(row["Total Fees"]) else 0)
                ),

                reference_number=text(row["Reference Num"]),
                invoice_number=text(row["Invoice Num"])
            )

            temp_html = output_dir / f"{row['Shareholder Name']}-{output_name}.html"
            pdf_path = output_dir / f"{row['Shareholder Name']}-{output_name}.pdf"

            temp_html.write_text(html_content, encoding="utf-8")

            page = browser.new_page()
            page.goto(temp_html.resolve().as_uri())
            page.pdf(
                path=str(pdf_path),
                format="A4",
                print_background=True
            )
            page.close()
            
            temp_html.unlink()

        browser.close()

def main():
    Tk().withdraw()

    excel_file = askopenfilename(
        title="Select Investor Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not excel_file:
        messagebox.showerror("Error", "No file selected.")
        sys.exit(1)
    else:    
        df = pd.read_excel(excel_file, sheet_name = 1)
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce")
    
    required_columns = [
        'Shareholder Details', 'Shareholder Name', 'Shareholder Address', 
        'Shareholder Email', 'SPV Name', 'SPV Reg Num', "SPV Director Name", 'Company Name',
        'Company Reg Num', 'Currency', 'Conversion Description', 'Date', 'Due Date', 
        'One Time Fees', 'Admin Service Fees', 'Miscellaneous Description', "Miscellaneous Fees",
        'Total Fees', 'Hard Commit Sum', 'Reference Num', 'Invoice Num'
    ]

    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise Exception(f"Missing columns: {missing}")

    env = Environment(loader=FileSystemLoader(resource_path("templates")))
    
    generate_capital_call = messagebox.askyesno(
        "Generate Capital Calls?",
        "Do you want to generate Capital Call documents as well?"
    )

    generate_documents(df, "invoice.html", "invoice", env)

    if generate_capital_call:
        generate_documents(df, "capital_call.html", "capital-call", env)

    messagebox.showinfo("Done", "Documents generated successfully.")

if __name__ == "__main__":
    main()