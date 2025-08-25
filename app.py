from flask import Flask, render_template, request, send_file
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from io import BytesIO
from datetime import datetime

app = Flask(__name__)

# Coordinate dei campi in mm (da regolare con griglia/Acrobat)
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
    # Prendi i valori dal form
    dati = {campo: request.form.get(campo, '') for campo in COORDS.keys()}

    # Imposta DATA_INTERVENTO = data odierna
    oggi = datetime.today().strftime("%d/%m/%Y")
    dati['DATA_INTERVENTO'] = oggi

    # PDF temporaneo con testo
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(210*mm, 297*mm))  # A4
    can.setFont("Helvetica", 12)

    # Scrivi testo nei punti indicati
    for campo, (x, y) in COORDS.items():
        can.drawString(x, y, dati[campo])

    can.save()
    packet.seek(0)

    # Carica modello PDF
    existing_pdf = PdfReader("modello.pdf")
    new_pdf = PdfReader(packet)
    output = PdfWriter()

    # Sovrapponi la pagina
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    # Salva PDF in memoria
    output_stream = BytesIO()
    output.write(output_stream)
    output_stream.seek(0)

    return send_file(output_stream, download_name="modulo_compilato.pdf", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)