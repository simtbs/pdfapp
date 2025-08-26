from flask import Flask, render_template, request, url_for
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from io import BytesIO
from datetime import datetime
from flask_mail import Mail, Message
import os
import re  # <-- aggiunto per pulire il nome file

app = Flask(__name__)

# ----------------------
# Configurazione Flask-Mail con variabili d'ambiente
# ----------------------
app.config['MAIL_SERVER'] = os.environ.get('MAIL_SERVER', 'smtp.sendgrid.net')
app.config['MAIL_PORT'] = int(os.environ.get('MAIL_PORT', 587))
app.config['MAIL_USE_TLS'] = os.environ.get('MAIL_USE_TLS', 'True') == 'True'
app.config['MAIL_USE_SSL'] = os.environ.get('MAIL_USE_SSL', 'False') == 'True'
app.config['MAIL_USERNAME'] = os.environ.get('MAIL_USERNAME', 'apikey')
app.config['MAIL_PASSWORD'] = os.environ.get('MAIL_PASSWORD', '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)
app.config['MAIL_DEFAULT_SENDER'] = os.environ.get('MAIL_DEFAULT_SENDER', 'sp.perniciaro@gmail.com')

mail = Mail(app)

# Coordinate dei campi
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
    # Recupero modello
    modello_path = os.path.join(os.getcwd(), "modello.pdf")
    if not os.path.exists(modello_path):
        return "Errore: 'modello.pdf' non trovato nella root del progetto.", 500

    os.makedirs("static", exist_ok=True)

    # Dati dal form
    dati = {campo: request.form.get(campo, '') for campo in COORDS.keys()}
    dati['DATA_INTERVENTO'] = datetime.today().strftime("%d/%m/%Y")

    # Layer testo
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

    # Nome file dinamico dal campo WR_IMPIANTO
    wr_valore = dati.get('WR_IMPIANTO', 'modulo')
    wr_valore_sicuro = re.sub(r'[^a-zA-Z0-9_-]', '_', wr_valore)
    filename = f"{wr_valore_sicuro}.pdf"

    # Salvataggio PDF
    file_abs_path = os.path.join("static", filename)
    with open(file_abs_path, "wb") as f:
        output.write(f)

    # URL di download
    file_url_abs = url_for('static', filename=filename, _external=True)

    # Invia email con allegato
    try:
        msg = Message("Report FTTH", recipients=["s.perniciaro@simt.it"])  # 🔴 cambia destinatario
        msg.body = "REPORT DELIVERYY FTTH"
        with open(file_abs_path, "rb") as f:
            msg.attach(filename, "application/pdf", f.read())
        mail.send(msg)
    except Exception as e:
        return f"Errore durante l'invio dell'email: {e}", 500

    # Pagina di successo
    return render_template("success.html", file_url=file_url_abs, filename=filename)

if __name__ == "__main__":
    app.run(debug=True)
