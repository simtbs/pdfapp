from flask import Flask, render_template, request, url_for
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from io import BytesIO
from datetime import datetime
from unidecode import unidecode
import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import base64

app = Flask(__name__)

# ----------------------
# Coordinate dei campi
# ----------------------
COORDS = {
    'TGU': (27*mm, 260*mm),
    'INDIRIZZO_CLIENTE': (48*mm, 243*mm),
    'COMUNE_INTERVENTO': (55*mm, 237*mm),
    'WR_IMPIANTO': (28*mm, 219*mm),
    'ATTIVITA_ESEGUITA': (27*mm, 207*mm),
    'PRODOTTO_1': (94*mm, 194*mm),
    'NMU_PRODOTTO_1': (31*mm, 188*mm),
    'PRODOTTO_2': (50*mm, 182*mm),
    'NMU_PRODOTTO_2': (31*mm, 176*mm),
    'PRODOTTO_3': (50*mm, 170*mm),
    'NMU_PRODOTTO_3': (31*mm, 164*mm),
    'PERCORSO_DI_RETE': (13*mm, 33*mm),
    'BORCHIA_OTTICA': (13*mm, 24*mm),
    'DATA_INTERVENTO': (86*mm, 33*mm),
    'CODICE_COLLAUDO': (13*mm, 10*mm)
}


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/genera', methods=['POST'])
def genera_pdf():
    modello_path = os.path.join(os.getcwd(), "modello.pdf")
    if not os.path.exists(modello_path):
        return "Errore: 'modello.pdf' non trovato.", 500

    os.makedirs("static", exist_ok=True)

    # Dati dal form
    dati = {campo: request.form.get(campo, '') for campo in COORDS.keys()}
    dati['DATA_INTERVENTO'] = datetime.today().strftime("%d/%m/%Y")

    # Layer testo PDF
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(210*mm, 297*mm))
    can.setFont("Helvetica", 12)
    for campo, (x, y) in COORDS.items():
        can.drawString(x, y, dati.get(campo, ''))
    can.save()
    packet.seek(0)

    # Merge PDF
    existing_pdf = PdfReader(modello_path)
    new_pdf = PdfReader(packet)
    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    # Nome PDF dinamico
    wr_safe = unidecode(dati.get('WR_IMPIANTO', 'modulo')).replace(" ", "_")
    filename = f"{wr_safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    file_abs_path = os.path.join("static", filename)
    with open(file_abs_path, "wb") as f:
        output.write(f)

    file_url_abs = url_for('static', filename=filename, _external=True)

    # ----------------------
    # Invio email con SendGrid API
    # ----------------------
    import os
import sendgrid
from sendgrid.helpers.mail import Mail

SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY")

@app.route('/genera', methods=['POST'])
def genera_pdf():
    ...
    try:
        sg = sendgrid.SendGridAPIClient(SENDGRID_API_KEY)
        message = Mail(
            from_email="s.perniciaro@simt.it",   # deve essere verificata su SendGrid
            to_emails="s.perniciaro@simt.it",
            subject="Report FTTH",
            html_content="<p>REPORT DELIVERY FTTH</p>"
        )
        with open(file_abs_path, "rb") as f:
            message.add_attachment(f.read(), "application/pdf", filename, "attachment")
        sg.send(message)
    except Exception as e:
        return f"Errore durante l'invio dell'email: {e}", 500
        
    return render_template("success.html", file_url=file_url_abs, filename=filename)


if __name__ == "__main__":
    app.run(debug=True)
