#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Labels Generator - Final Complete Version
Leest CSV-bestand, genereert verzendlabels tekstbestand en maakt direct PDF output.
Gebruikt: /Users/timlind/Desktop/orders sales, oprecht en ontspannen.csv
Output: enkel PDF bestand
"""

import csv
import sys
import os
from collections import OrderedDict
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageTemplate, Frame
from reportlab.lib import colors
from reportlab.lib.units import mm
from PyPDF2 import PdfMerger

# ------------------------------
# CSV LEES EN FILTER FUNCTIES
# ------------------------------

def read_csv_file(filename):
    """Lees het CSV-bestand en retourneer de data."""
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            return list(reader)
    except FileNotFoundError:
        print(f"Fout: Bestand '{filename}' niet gevonden.")
        return None
    except Exception as e:
        print(f"Fout bij het lezen van het bestand: {e}")
        return None

def format_name(row):
    """Formatteer de naam op basis van beschikbare velden."""
    company = row['company'].strip() if row['company'] and row['company'].strip() else ''
    firstname = row['firstname'].strip() if row['firstname'] and row['firstname'].strip() else ''
    lastname = row['lastname'].strip() if row['lastname'] and row['lastname'].strip() else ''

    if company:
        # Als er een bedrijf is, gebruik bedrijfsnaam
        name = company
        # Voeg persoonlijke naam toe als die bestaat en niet overeenkomt met bedrijfsnaam
        if firstname and lastname:
            personal_name = f"{firstname} {lastname}"
            if personal_name.lower() not in company.lower():
                name = f"{company} - {personal_name}"
    else:
        # Geen bedrijf, gebruik persoonlijke naam
        name = f"{firstname} {lastname}".strip()

    return name

def format_address(row):
    """Formatteer het adres."""
    street = row['street'].strip() if row['street'] and row['street'].strip() else ''
    housenumber = row['housenumber'].strip() if row['housenumber'] and row['housenumber'].strip() else ''
    suffix = row['housenumber_suffix'].strip() if row['housenumber_suffix'] and row['housenumber_suffix'].strip() else ''

    # Combineer huisnummer met toevoeging
    full_housenumber = housenumber
    if suffix:
        full_housenumber += suffix

    # Combineer straat en huisnummer
    address = f"{street} {full_housenumber}".strip()

    return address

def format_postal(row):
    """Formatteer postcode en plaats."""
    zipcode = row['zipcode'].strip() if row['zipcode'] and row['zipcode'].strip() else ''
    city = row['city'].strip() if row['city'] and row['city'].strip() else ''

    return f"{zipcode} {city}".strip()

def should_include_product(product):
    """Bepaal of het product overeenkomt met een van de toegestane producten."""
    if not product:
        return False

    # Product moet gelijk zijn aan 'Boek: Sales, oprecht en ontspannen' of 'Boek Oprecht en Ontspannen Sales'
    allowed_products = [
        "Boek: Sales, oprecht en ontspannen",
        "Boek Oprecht en Ontspannen Sales"
    ]

    return product.strip() in allowed_products

def generate_shipping_labels(data):
    """Genereer verzendlabels van de CSV data."""
    # Sorteer data op 'paid_at' (datum/tijd van betaling)
    sorted_data = sorted(data, key=lambda row: row['paid_at'] if row['paid_at'] else '')

    unique_addresses = OrderedDict()

    for row in sorted_data:
        # Filter op toegestane producten
        if not should_include_product(row['product']):
            continue

        # Maak een unieke sleutel voor het adres
        name = format_name(row)
        address = format_address(row)
        postal = format_postal(row)

        address_key = f"{name}|{address}|{postal}"

        # Voeg toe aan unieke adressen (behaalt de sortering volgens paid_at)
        if address_key not in unique_addresses:
            unique_addresses[address_key] = {
                'name': name,
                'address': address,
                'postal': postal
            }

    # Genereer de labels voor PDF (geen tekstbestand meer)
    labels = []
    for address_data in unique_addresses.values():
        label_text = f"{address_data['name']}\n{address_data['address']}\n{address_data['postal']}"
        labels.append(label_text)

    return labels

# ------------------------------
# PDF GENERATIE FUNCTIES
# ------------------------------

def create_table_with_labels(labels, start_index):
    """Maak een 8x3 tabel met labels vanaf start_index."""

    # A4 afmetingen: 210mm x 297mm
    # 8 rijen en 3 kolommen - exacte berekening voor volledige pagina vulling
    col_widths = [70*mm, 70*mm, 70*mm]    # 210mm / 3 = 70mm per kolom
    row_heights = [37.125*mm] * 8        # 297mm / 8 = 37.125mm per rij

    # Maak tabel data (8 rijen, 3 kolommen) met echte labels
    table_data = []

    for i in range(8):  # 8 rijen
        row = []
        for j in range(3):  # 3 kolommen
            label_index = start_index + (i * 3) + j

            if label_index < len(labels):
                # Gebruik het echte label
                label_text = labels[label_index]
                row.append(label_text)
            else:
                # Geen label meer beschikbaar
                row.append('')

        table_data.append(row)

    # Maak de tabel met GEEN extra ruimte of padding
    table = Table(
        table_data,
        colWidths=col_widths,
        rowHeights=row_heights
    )

    # Stijl de tabel met GEEN padding
    table.setStyle(TableStyle([
        # Cell borders
        ('GRID', (0, 0), (-1, -1), 2, colors.black),

        # Tekst centrering
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

        # Tekst eigenschappen
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('LEADING', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),

        # ABSOLUUT GEEN padding of marges
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),

        # Geen extra spacing
        ('NOSPLIT', (0, 0), (-1, -1)),

        # Tabel niveau instellingen
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.white for _ in range(8)]),
    ]))

    return table

def create_pdf_from_labels(labels, output_file):
    """Maak het volledige PDF document met alle labels."""

    # Bereken hoeveel pagina's nodig zijn (24 labels per pagina: 8×3)
    labels_per_page = 8 * 3
    pages_needed = (len(labels) + labels_per_page - 1) // labels_per_page

    # Maak alle tabellen
    tables = []
    for page_num in range(pages_needed):
        start_index = page_num * labels_per_page
        table = create_table_with_labels(labels, start_index)
        tables.append(table)

    print(f"PDF document wordt {pages_needed} pagina('s).")
    print(f"Afmetingen: A4 (210×297mm) met 8×3 tabel per pagina.")
    print(f"Elke cel: 70×37.125mm (gehele pagina dekking)")

    # Maak individuele PDF's en combineer ze
    pdf_pages = []

    for i, table in enumerate(tables):
        # Maak een tijdelijk document voor elke pagina
        temp_doc = SimpleDocTemplate(
            output_file.replace('.pdf', f'_temp_{i}.pdf'),
            pagesize=A4,
            leftMargin=0,
            rightMargin=0,
            topMargin=0,
            bottomMargin=0,
            showBoundary=0
        )

        def on_page(canvas, doc):
            canvas.saveState()
            canvas.resetTransforms()
            table.wrapOn(canvas, A4[0], A4[1])
            table.drawOn(canvas, 0, 0)
            canvas.restoreState()

        frame = Frame(0, 0, A4[0], A4[1], leftPadding=0, rightPadding=0,
                      topPadding=0, bottomPadding=0)

        template = PageTemplate(id=f'page_{i}', frames=[frame], onPage=on_page)
        temp_doc.addPageTemplates([template])

        # Gebruik een dummy flowable
        temp_doc.build([table])

    # Combineer alle PDF's
    merger = PdfMerger()

    for i in range(len(tables)):
        temp_file = output_file.replace('.pdf', f'_temp_{i}.pdf')
        merger.append(temp_file)

    merger.write(output_file)
    merger.close()

    # Verwijder tijdelijke bestanden
    for i in range(len(tables)):
        temp_file = output_file.replace('.pdf', f'_temp_{i}.pdf')
        if os.path.exists(temp_file):
            os.remove(temp_file)

    print(f"PDF succesvol aangemaakt: {output_file}")

# ------------------------------
# HOOFDFUNCTIE
# ------------------------------

def main():
    """Hoofdfunctie - complete proces van CSV naar PDF."""

    # Input en output bestanden
    input_file = "orders sales, oprecht en ontspannen.csv"
    output_file = "verzendlabels.pdf"

    print("Labels Generator - Complete CSV naar PDF")
    print("=" * 50)
    print(f"Input CSV: {input_file}")
    print(f"Output PDF: {output_file}")

    # Controleer of het input bestand bestaat
    if not os.path.exists(input_file):
        print(f"Fout: Input bestand '{input_file}' niet gevonden.")
        return

    # Lees de data uit CSV
    print("Lezen van CSV bestand...")
    data = read_csv_file(input_file)
    if not data:
        print("Kan het CSV-bestand niet lezen. Script afgebroken.")
        return

    print(f"{len(data)} rijen gevonden in het CSV-bestand.")

    # Genereer de labels met filtering
    print("Genereren van verzendlabels met product filtering...")
    labels = generate_shipping_labels(data)

    if not labels:
        print("Geen geldige labels gevonden na filtering. Script afgebroken.")
        return

    print(f"{len(labels)} unieke verzendlabels gegenereerd.")

    # Genereer de PDF direct
    print("Aanmaken van PDF document...")
    create_pdf_from_labels(labels, output_file)

    print(f"\n✅ KLAAR! {len(labels)} verzendlabels verwerkt in PDF bestand.")
    print(f"📄 Output: {output_file}")
    print("🖨️  PDF is klaar om af te drukken.")
    print("📋 Elke pagina bevat 8×3 labels die de hele A4 pagina vullen.")

if __name__ == "__main__":
    main()