import os
import re
import csv
from io import StringIO
import PyPDF2
from datetime import datetime


def format_date(date_str):
    # Intenta analizar según el formato aaaa-mm-dd
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%d/%m/%Y")  # Formatea a dd/mm/aaaa
    except ValueError:
        # Retorna la fecha original si ya está en el formato dd/mm/aaaa o hay un error
        return date_str


def load_provider_data(filename):
    with open(filename, newline='') as csvfile:
        reader = csv.reader(csvfile)
        next(reader, None)  # Skip header if present
        return {rows[1]: rows[0] for rows in reader}


def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        return ''.join(page.extract_text() for page in reader.pages)


def get_matches(document, patterns):
    return {key: re.search(pattern, document, re.DOTALL | re.IGNORECASE) for key, pattern in patterns.items()}


def format_serie_numero(serie_numero):
    return re.sub(r'FA(\d{2}) Nº (\d{8})', r'FA\1-\2', serie_numero) if serie_numero else None


def find_destinatario(document, destinatarios):
    for key, value in destinatarios.items():
        if re.search(key, document, re.IGNORECASE):
            return value
    # Si no se encuentra, busca usando la expresión regular ajustada para "CLIENTE"
    match = re.search(r'CLIENTE\s*:?[\s]*([^\n]+)', document, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    # Otra búsqueda general si el formato específico falla
    general_match = re.search(r'SEÑOR\s+(.*?):|Señor\(es\)\s*:\s*(.*?)\n', document, re.IGNORECASE)
    if general_match:
        # Asegurarse de devolver el grupo adecuado que no sea None
        return (general_match.group(1) or general_match.group(2)).strip()

    return None


def extract_invoice_info(document, providers, destinatarios):
    patterns = {
        'ruc_emisor': r'(?<=\D)(20\d{9}|10\d{9})',
        'serie_numero': r'(FA\d{2} Nº \d{8}|F\d{3}-\d{8}|E\d{3}-\d{4}|E\d{3}-\d{1,4}|E\d{1,3}-\d{1,8}\s*|E001-\s*\d{1,8}|Nro\.\s*?F\d{3}-\d{8}|F\d{3}-\d+)',
        'fecha_emision': r'(\d{2}/\d{2}/\d{4})|(\d{4}-\d{2}-\d{2})',
        'monto_total': r'TOTAL\s+IMPORTE\s+VENTA\s*:\s*S/\s*([\d,]+\.\d{2})|IMPORTE\s+TOTAL\s*:\s*S/\s*([\d,]+\.\d{2})|CIENTO.*?(\d+\.\d{2})|Importe\s+Total\s*:\s*S/\s*([\d,]+\.\d{2})|Importe\s+Total\s*S/\s*([\d,]+\.\d{2})|Importe\s+Total\s+S/\s*(\d+\.\d{2})|Importe\s+Total\s*?(\d+\.\d{2})|IMPORTE\s+TOTAL\s*?(\d+\.\d{2})'
    }

    matches = get_matches(document, patterns)
    fecha_emision_match = matches['fecha_emision']
    # Determinar cuál grupo capturó la fecha si hay una coincidencia
    if fecha_emision_match:
        fecha_emision = fecha_emision_match.group(1) if fecha_emision_match.group(1) else fecha_emision_match.group(2)
        fecha_emision = format_date(fecha_emision)  # Formatear si es necesario
    else:
        fecha_emision = None
    ruc_emisor = matches['ruc_emisor'].group(0) if matches['ruc_emisor'] else None
    empresa_emisora = providers.get(ruc_emisor, 'Desconocido') if ruc_emisor else 'Desconocido'
    serie_numero = format_serie_numero(matches['serie_numero'].group(0) if matches['serie_numero'] else None)
    monto_total = next((float(match.replace(',', '')) for match in
                        [matches['monto_total'].group(i) for i in range(1, 9) if
                         matches['monto_total'] and matches['monto_total'].group(i)]), None)
    empresa_destinataria = find_destinatario(document, destinatarios)

    return [empresa_emisora, ruc_emisor, serie_numero, monto_total, fecha_emision, empresa_destinataria]


def process_invoices(directory, providers, destinatarios):
    invoices = [extract_invoice_info(extract_text_from_pdf(os.path.join(directory, filename)), providers, destinatarios)
                for filename in os.listdir(directory) if filename.endswith('.pdf')]
    headers = ["PROVEEDOR", "RUC", "NR-SERIE", "MONTO", "FECHA EMITIDA", "EMPRESA DESTINATARIA"]
    output = StringIO()
    csv.writer(output).writerow(headers)
    csv.writer(output).writerows(invoices)
    return output.getvalue()


def save_csv(data, filename):
    with open(filename, 'w', newline='') as file:
        file.write(data)
    print(f"El archivo CSV se ha guardado como: {filename}")


if __name__ == "__main__":
    providers_filename = 'proveedores.csv'
    providers = load_provider_data(providers_filename)
    directory = 'facturasADigitalizar'
    destinatarios = {
        'AYG NUVA SAC': 'AYG Nuva',
        'AYG NUVA S.A.C.': 'AYG Nuva',
        'AYG RICOS SAC': 'AYG Ricos',
        'AYG RICOS S.A.C.': 'AYG Ricos',
        'NUVA SERVICE SAC': 'Nuva Service',
        'NUVA SERVICE S.A.C.': 'Nuva Service',
        'NUVA PROCESOS ALIMENTARIOS SAC': 'Nuva Procesos Alimentarios',
        'NUVA PROCESOS ALIMENTARIOS S.A.C.': 'Nuva Procesos Alimentarios',
        'NUVA KALLPA EIRL': 'Nuva Kallpa',
        'NUVA	KALLPA	E.I.R.L': 'Nuva Kallpa',
        'NUVA    KALLPA    E.I.R.L': 'Nuva Kallpa',
        "A	Y	G	RICO'S	S.A.C.": 'AYG Ricos',
        'A Y G RICO\'S SAC': 'AYG Ricos',
        'A Y G RICO S S.A.C.': 'AYG Ricos',
        "A Y G RICO'S S.A.C.": 'AYG Ricos',
        "A Y G RICO´S S.A.C.": 'AYG Ricos',
        "A Y G RICO`S S.A.C.": 'AYG Ricos'
    }
    csv_data = process_invoices(directory, providers, destinatarios)
    save_csv(csv_data, 'facturas.csv')
