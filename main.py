import os, csv, json, shutil, tempfile
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename
from docx import Document
import docx2pdf
from PyPDF2 import PdfMerger
import comtypes.client  # For fallback conversion using COM automation
from io import BytesIO
import pythoncom       # For COM initialization

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Required for flash messaging

# Base directory for storing uploaded scorecard templates (separate from HTML templates)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
SCTEMP_DIR = os.path.join(BASE_DIR, 'SCTEMP')
if not os.path.exists(SCTEMP_DIR):
    os.makedirs(SCTEMP_DIR)

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

# ===================== Document Processing Functions =====================

def merge_runs_in_paragraph(paragraph):
    if len(paragraph.runs) <= 1:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    for run in paragraph.runs:
        run.text = ""
    paragraph.runs[0].text = full_text

def replace_text_in_paragraph(paragraph, placeholder, replacement):
    if placeholder in paragraph.text:
        new_text = paragraph.text.replace(placeholder, replacement)
        paragraph.runs[0].text = new_text

def replace_placeholders_in_paragraphs(paragraphs, placeholders):
    for paragraph in paragraphs:
        merge_runs_in_paragraph(paragraph)
        for placeholder, replacement in placeholders:
            replace_text_in_paragraph(paragraph, placeholder, replacement)

def replace_text_in_doc(doc, rows, mapping, cards_per_page=4):
    """
    Replace placeholders in the document.
    For each scorecard (from 1 to cards_per_page) and for each CSV header mapping,
    the function looks for placeholders like "PLACEHOLDER_1", "PLACEHOLDER_2", etc.
    """
    all_placeholders = []
    for i in range(1, cards_per_page+1):
        row = rows[i-1] if i-1 < len(rows) else None
        for csv_header, placeholder in mapping.items():
            value = row.get(csv_header, "") if row else ""
            all_placeholders.append((f"{placeholder}_{i}", value))
    replace_placeholders_in_paragraphs(doc.paragraphs, all_placeholders)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders_in_paragraphs(cell.paragraphs, all_placeholders)

def merge_two_pdfs(pdf1, pdf2, merged_pdf):
    merger = PdfMerger()
    merger.append(pdf1)
    merger.append(pdf2)
    merger.write(merged_pdf)
    merger.close()

def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Convert a DOCX file to PDF.
    First attempts to use docx2pdf.convert; if that fails,
    falls back to using COM automation (Windows only).
    """
    try:
        docx2pdf.convert(docx_path, pdf_path)
    except Exception as e:
        # Fallback conversion using COM automation with explicit COM initialization.
        pythoncom.CoInitialize()
        wdFormatPDF = 17
        try:
            word = comtypes.client.CreateObject("Word.Application")
        except Exception as e2:
            pythoncom.CoUninitialize()
            raise e2
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        try:
            doc.SaveAs2(pdf_path, FileFormat=wdFormatPDF)
        except Exception as e3:
            doc.Close()
            word.Quit()
            pythoncom.CoUninitialize()
            raise e3
        doc.Close()
        word.Quit()
        pythoncom.CoUninitialize()

def generate_scorecard(template_path, csv_path, mapping, cards_per_page=4, back_pdf_path=None, temp_dir=None):
    """
    Generate the final scorecard PDF.
    All temporary files (e.g. intermediate DOCX/PDFs) are created in temp_dir.
    """
    final_pdf_list = []
    with open(csv_path, newline='', encoding='latin-1') as f:
        sample = f.read(1024)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample)
        reader = csv.DictReader(f, dialect=dialect)
        rows = [row for row in reader if any(value.strip() for value in row.values())]
    for i in range(0, len(rows), cards_per_page):
        group = rows[i:i+cards_per_page]
        front_doc = Document(template_path)
        replace_text_in_doc(front_doc, group, mapping, cards_per_page=cards_per_page)
        temp_front_docx = os.path.join(temp_dir, f"temp_front_{i//cards_per_page}.docx")
        temp_front_pdf = os.path.join(temp_dir, f"temp_front_{i//cards_per_page}.pdf")
        page_pdf = os.path.join(temp_dir, f"page_{i//cards_per_page}.pdf")
        front_doc.save(temp_front_docx)
        convert_docx_to_pdf(temp_front_docx, temp_front_pdf)
        if back_pdf_path and os.path.exists(back_pdf_path):
            merge_two_pdfs(temp_front_pdf, back_pdf_path, page_pdf)
        else:
            shutil.copy(temp_front_pdf, page_pdf)
        final_pdf_list.append(page_pdf)
        os.remove(temp_front_docx)
        os.remove(temp_front_pdf)
    output_pdf = os.path.join(temp_dir, "Final_Scorecards.pdf")
    merger = PdfMerger()
    sorted_pdf_list = sorted(final_pdf_list, key=lambda x: int(x.split('_')[-1].split('.')[0]))
    for pdf in sorted_pdf_list:
        merger.append(pdf)
    merger.write(output_pdf)
    merger.close()
    for pdf in final_pdf_list:
        os.remove(pdf)
    return output_pdf

# ===================== Flask Routes =====================

@app.route('/')
def index():
    """List all available scorecard templates."""
    templates_list = []
    for sport in os.listdir(SCTEMP_DIR):
        sport_dir = os.path.join(SCTEMP_DIR, sport)
        if os.path.isdir(sport_dir):
            for template_name in os.listdir(sport_dir):
                template_dir = os.path.join(sport_dir, template_name)
                if os.path.isdir(template_dir):
                    templates_list.append({'sport': sport, 'template': template_name})
    return render_template('index.html', templates=templates_list)

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    """
    Upload a new scorecard template.
    The user must provide:
     - Sport and Template Name
     - A Word template (front) (.docx)
     - A CSV template file for data entry
     - Optionally, a static back design (.pdf)
    These files are stored under SCTEMP/(SPORT)/(TEMPLATENAME)/.
    """
    if request.method == 'POST':
        sport = request.form.get('sport')
        template_name = request.form.get('template_name')
        if not sport or not template_name:
            flash("Sport and Template Name are required.")
            return redirect(request.url)
        sport = secure_filename(sport)
        template_name = secure_filename(template_name)
        template_dir = os.path.join(SCTEMP_DIR, sport, template_name)
        os.makedirs(template_dir, exist_ok=True)
        # Save the front Word template
        front_file = request.files.get('front_file')
        if not front_file or not allowed_file(front_file.filename, {'docx'}):
            flash("A valid Word template file (.docx) is required.")
            return redirect(request.url)
        front_path = os.path.join(template_dir, 'template_front.docx')
        front_file.save(front_path)
        # Save the CSV file (template for data entry)
        csv_file = request.files.get('csv_file')
        if not csv_file or not allowed_file(csv_file.filename, {'csv'}):
            flash("A valid CSV file is required.")
            return redirect(request.url)
        csv_path = os.path.join(template_dir, 'template_data.csv')
        csv_file.save(csv_path)
        # Optionally save a static back design
        back_option = request.form.get('back_option')
        if back_option == 'yes':
            back_file = request.files.get('back_file')
            if back_file and allowed_file(back_file.filename, {'pdf'}):
                back_path = os.path.join(template_dir, 'template_back.pdf')
                back_file.save(back_path)
            else:
                flash("Back design selected but no valid PDF uploaded.")
                return redirect(request.url)
        return redirect(url_for('mapping', sport=sport, template_name=template_name))
    return render_template('upload.html')

@app.route('/mapping/<sport>/<template_name>', methods=['GET', 'POST'])
def mapping(sport, template_name):
    """
    Read the CSV headers from the uploaded CSV template and prompt the user to map
    each header to a placeholder text. Also ask how many scorecards per page will be printed.
    The mapping (with cards_per_page) is stored as JSON.
    """
    template_dir = os.path.join(SCTEMP_DIR, secure_filename(sport), secure_filename(template_name))
    csv_path = os.path.join(template_dir, 'template_data.csv')
    if not os.path.exists(csv_path):
        flash("CSV file not found for this template.")
        return redirect(url_for('index'))
    with open(csv_path, newline='', encoding='latin-1') as f:
        sample = f.read(1024)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample)
        reader = csv.reader(f, dialect=dialect)
        headers = next(reader)
    if request.method == 'POST':
        mapping_data = {}
        for header in headers:
            placeholder = request.form.get(header)
            mapping_data[header] = placeholder if placeholder else header
        try:
            cards_per_page = int(request.form.get('cards_per_page', 4))
        except ValueError:
            cards_per_page = 4
        mapping_file = os.path.join(template_dir, 'mapping.json')
        with open(mapping_file, 'w') as f:
            json.dump({"cards_per_page": cards_per_page, "mapping": mapping_data}, f)
        flash("Mapping saved successfully.")
        return redirect(url_for('index'))
    instructions = ("For each phrase to be replaced on the template, append an underscore and a number "
                    "corresponding to the scorecard position on the page (e.g., if 3 scorecards per page, "
                    "use _1, _2, _3).")
    return render_template('mapping.html', headers=headers, sport=sport, template_name=template_name, instructions=instructions)

@app.route('/generate/<sport>/<template_name>', methods=['GET', 'POST'])
def generate(sport, template_name):
    """
    GET: Display a page to download the CSV template so the user can enter their data,
         and a form to upload the completed CSV.
    POST: Accept the filled CSV, then create a temporary working directory where all files
          (intermediate and final) are processed. The final PDF is then returned (read into memory)
          and the temporary directory is purged. A cookie is set so the client-side JavaScript
          can detect the completion and redirect the user.
    """
    template_dir = os.path.join(SCTEMP_DIR, secure_filename(sport), secure_filename(template_name))
    front_path = os.path.join(template_dir, 'template_front.docx')
    mapping_file = os.path.join(template_dir, 'mapping.json')
    back_path = os.path.join(template_dir, 'template_back.pdf')
    if request.method == 'POST':
        filled_csv_file = request.files.get('filled_csv')
        if not filled_csv_file or not allowed_file(filled_csv_file.filename, {'csv'}):
            flash("Please upload a valid CSV file with your data.")
            return redirect(request.url)
        with tempfile.TemporaryDirectory() as temp_dir:
            filled_csv_path = os.path.join(temp_dir, 'filled_data.csv')
            filled_csv_file.save(filled_csv_path)
            with open(mapping_file, 'r') as f:
                mapping_json = json.load(f)
            cards_per_page = mapping_json.get("cards_per_page", 4)
            mapping_data = mapping_json.get("mapping", {})
            output_pdf = generate_scorecard(front_path, filled_csv_path, mapping_data,
                                            cards_per_page=cards_per_page,
                                            back_pdf_path=back_path if os.path.exists(back_path) else None,
                                            temp_dir=temp_dir)
            with open(output_pdf, 'rb') as pdf_file:
                pdf_bytes = pdf_file.read()
            from flask import make_response
            response = make_response(send_file(BytesIO(pdf_bytes),
                             download_name=os.path.basename(output_pdf),
                             mimetype='application/pdf',
                             as_attachment=True))
            response.set_cookie("fileDownload", "true", max_age=60)
            return response
    else:
        return render_template('generate.html', sport=sport, template_name=template_name)

@app.route('/download_csv/<sport>/<template_name>')
def download_csv(sport, template_name):
    """
    Serve the original CSV template file so the user can download and fill it with data.
    """
    template_dir = os.path.join(SCTEMP_DIR, secure_filename(sport), secure_filename(template_name))
    csv_template_path = os.path.join(template_dir, 'template_data.csv')
    if not os.path.exists(csv_template_path):
         flash("CSV template not found.")
         return redirect(url_for('index'))
    return send_file(csv_template_path, as_attachment=True)

@app.route('/delete/<sport>/<template_name>', methods=['POST'])
def delete_template(sport, template_name):
    """
    Delete the entire directory for a given scorecard template.
    """
    template_dir = os.path.join(SCTEMP_DIR, secure_filename(sport), secure_filename(template_name))
    if os.path.exists(template_dir):
        shutil.rmtree(template_dir)
        flash("Template deleted successfully.")
    else:
        flash("Template not found.")
    return redirect(url_for('index'))

@app.route('/about')
def about():
    return render_template('about.html')

if __name__ == '__main__':
    app.run(debug=True)
