#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit App - Verzendlabels Generator
Upload CSV bestand en genereer verzendlabels als PDF
"""

import streamlit as st
import pandas as pd
import io
import tempfile
import os
from datetime import datetime, timedelta
from collections import OrderedDict
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageTemplate, Frame
from reportlab.lib import colors
from reportlab.lib.units import mm
from PyPDF2 import PdfMerger

# ------------------------------
# FUNCTIES UIT ORIGINELE SCRIPT
# ------------------------------

def read_csv_data(file):
    """Lees het CSV-bestand en retourneer de data als pandas DataFrame."""
    try:
        # Probeer CSV te lezen
        df = pd.read_csv(file)
        return df
    except Exception as e:
        st.error(f"Fout bij het lezen van het CSV-bestand: {e}")
        return None

def format_name(row):
    """Formatteer de naam op basis van beschikbare velden."""
    company = str(row['company']).strip() if pd.notna(row['company']) and str(row['company']).strip() else ''
    firstname = str(row['firstname']).strip() if pd.notna(row['firstname']) and str(row['firstname']).strip() else ''
    lastname = str(row['lastname']).strip() if pd.notna(row['lastname']) and str(row['lastname']).strip() else ''

    if company and company != 'nan':
        # Als er een bedrijf is, gebruik bedrijfsnaam
        name = company
        # Voeg persoonlijke naam toe op nieuwe regel als die bestaat en niet overeenkomt met bedrijfsnaam
        if firstname and lastname:
            personal_name = f"{firstname} {lastname}"
            if personal_name.lower() not in company.lower():
                name = f"{company}\n{personal_name}"
    else:
        # Geen bedrijf, gebruik persoonlijke naam
        name = f"{firstname} {lastname}".strip()

    return name

def format_address(row):
    """Formatteer het adres."""
    street = str(row['street']).strip() if pd.notna(row['street']) and str(row['street']).strip() else ''

    # Converteer huisnummer naar string en verwijder decimalen
    housenumber_raw = row['housenumber']
    if pd.notna(housenumber_raw):
        if isinstance(housenumber_raw, (int, float)):
            housenumber = str(int(housenumber_raw)) if housenumber_raw == int(housenumber_raw) else str(housenumber_raw).rstrip('.0')
        else:
            housenumber = str(housenumber_raw).strip()
    else:
        housenumber = ''

    suffix = str(row['housenumber_suffix']).strip() if pd.notna(row['housenumber_suffix']) and str(row['housenumber_suffix']).strip() and str(row['housenumber_suffix']).strip() != 'nan' else ''

    # Combineer huisnummer met toevoeging
    full_housenumber = housenumber
    if suffix:
        full_housenumber += suffix

    # Combineer straat en huisnummer
    address = f"{street} {full_housenumber}".strip()

    return address

def format_postal(row):
    """Formatteer postcode en plaats."""
    zipcode = str(row['zipcode']).strip() if pd.notna(row['zipcode']) and str(row['zipcode']).strip() else ''
    city = str(row['city']).strip() if pd.notna(row['city']) and str(row['city']).strip() else ''

    return f"{zipcode} {city}".strip()

def should_include_product(product, allowed_products):
    """Bepaal of het product overeenkomt met een van de toegestane producten."""
    if not product or product == 'nan':
        return False

    return product.strip() in allowed_products

def truncate_text_for_cell(text, max_chars_per_line=40, max_lines=4):
    """Breek tekst af om binnen een cel te passen, maar knip nooit volledige adressen af."""
    if not text:
        return text

    lines = text.split('\n')
    result_lines = []

    for line in lines:
        if len(line) <= max_chars_per_line:
            result_lines.append(line)
        else:
            # Breek lange lijnen af op woorden, maar bewaar alle tekst
            words = line.split(' ')
            current_line = []
            current_length = 0

            for word in words:
                if current_length + len(word) + (1 if current_line else 0) <= max_chars_per_line:
                    current_line.append(word)
                    current_length += len(word) + (1 if current_line else 0)
                else:
                    if current_line:
                        result_lines.append(' '.join(current_line))
                    current_line = [word]
                    current_length = len(word)

            if current_line:
                result_lines.append(' '.join(current_line))

    # Beperk aantal regels maar alleen als er echt te veel is (bijv. > 6 regels)
    if len(result_lines) <= 6:
        return '\n'.join(result_lines)
    else:
        return '\n'.join(result_lines[:6])  # Knip alleen af bij extreem lange tekst

def generate_shipping_labels(df, allowed_products, sort_order='newest_first', start_date=None, end_date=None):
    """Genereer verzendlabels van de CSV data met filters."""
    # Zet DataFrame om naar lijst van dictionaries
    data = df.to_dict('records')

    # Filter op datumbereik (indien opgegeven)
    if start_date or end_date:
        filtered_data = []
        for row in data:
            paid_at = row.get('paid_at', '')
            if pd.isna(paid_at) or not paid_at:
                continue

            # Probeer de datum te parsen
            try:
                # Probeer verschillende datumformaten
                date_formats = ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%m-%Y %H:%M:%S', '%d-%m-%Y']
                parsed_date = None
                for fmt in date_formats:
                    try:
                        parsed_date = datetime.strptime(str(paid_at), fmt)
                        break
                    except ValueError:
                        continue

                if parsed_date:
                    # Check if date is within range
                    if start_date and parsed_date < start_date:
                        continue
                    if end_date and parsed_date > end_date:
                        continue

                    # Update paid_at with parsed datetime for sorting
                    row['paid_at_parsed'] = parsed_date
                    filtered_data.append(row)
            except:
                # Als datum niet te parsen is, skip deze rij
                continue
        data = filtered_data
    else:
        # Geen datums filter, gebruik originele paid_at
        for row in data:
            # Probeer paid_at te parsen voor sortering
            paid_at = row.get('paid_at', '')
            try:
                date_formats = ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%m-%Y %H:%M:%S', '%d-%m-%Y']
                for fmt in date_formats:
                    try:
                        row['paid_at_parsed'] = datetime.strptime(str(paid_at), fmt)
                        break
                    except ValueError:
                        continue
            except:
                row['paid_at_parsed'] = datetime.min  # Geen datum = minimum voor sortering

    # Sorteer data op 'paid_at' op basis van sort_order
    reverse_order = (sort_order == 'newest_first')
    sorted_data = sorted(data, key=lambda row: row.get('paid_at_parsed', datetime.min), reverse=reverse_order)

    unique_addresses = OrderedDict()

    for row in sorted_data:
        # Filter op toegestane producten
        if not should_include_product(row['product'], allowed_products):
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

    # Genereer de labels voor PDF
    labels = []
    for address_data in unique_addresses.values():
        label_text = f"{address_data['name']}\n{address_data['address']}\n{address_data['postal']}"
        labels.append(label_text)

    return labels

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
                # Gebruik het echte label met tekst truncatie voor optimale celgrootte
                label_text = truncate_text_for_cell(labels[label_index])
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
        # Cell borders (transparant voor sticker vellen)
        ('GRID', (0, 0), (-1, -1), 2, colors.white),

        # Tekst centrering
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

        # Tekst eigenschappen - geoptimaliseerd voor 70mm x 37.125mm cellen
        ('FONTSIZE', (0, 0), (-1, -1), 10),  # Verder verhoogd voor betere leesbaarheid
        ('LEADING', (0, 0), (-1, -1), 12),  # 1.2x font size voor optimale regelspatiëring
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

    return output_file

# ------------------------------
# STREAMLIT UI
# ------------------------------

def main():
    """Hoofdfunctie voor Streamlit app."""
    st.set_page_config(
        page_title="Verzendlabels Generator",
        page_icon="",
        layout="centered"
    )

    st.title("Verzendlabels Generator")
    st.markdown("Upload een CSV-bestand met orders en genereer verzendlabels als PDF.")

    # File uploader
    uploaded_file = st.file_uploader(
        "Upload CSV-bestand",
        type=['csv'],
        help="Upload een CSV-bestand met kolommen: company, firstname, lastname, street, housenumber, housenumber_suffix, zipcode, city, product, paid_at"
    )

    if uploaded_file is not None:
        # Toon bestandsinformatie
        st.success(f"Bestand geüpload: {uploaded_file.name}")

        # Lees de data
        with st.spinner("CSV-bestand wordt gelezen..."):
            df = read_csv_data(uploaded_file)

        if df is not None:
            st.info(f"{len(df)} rijen geladen uit het CSV-bestand")

            # Filter en sorteer opties
            st.subheader("Filter en Sorteer Opties")

            # Sorteeroptie
            sort_order = st.radio(
                "Sorteer volgens betaaldatum:",
                options=["newest_first", "oldest_first"],
                format_func=lambda x: "Nieuwste eerst" if x == "newest_first" else "Oudste eerst",
                help="Sorteer de labels op basis van de betaaldatum"
            )

            # Haal datums uit paid_at kolom voor defaults
            df_dates = pd.to_datetime(df['paid_at'], errors='coerce').dropna()
            if not df_dates.empty:
                min_date = df_dates.min().date()
                max_date = df_dates.max().date()
            else:
                min_date = datetime.now().date() - timedelta(days=30)
                max_date = datetime.now().date()

            # Datumbereik filter
            st.subheader("Datumbereik (paid_at)")
            col_date1, col_date2 = st.columns(2)

            with col_date1:
                start_date = st.date_input(
                    "Vanaf datum:",
                    value=min_date,
                    help="Filter orders vanaf deze datum"
                )

            with col_date2:
                end_date = st.date_input(
                    "Tot en met datum:",
                    value=max_date,
                    help="Filter orders tot en met deze datum"
                )

            # Product filter
            st.subheader("Producten Filteren")

            # Haal unieke producten op
            unique_products = df['product'].dropna().unique().tolist()
            unique_products = [str(p) for p in unique_products if p and str(p) != 'nan']

            # Gebruik session state voor product selectie
            if 'product_selections' not in st.session_state:
                st.session_state.product_selections = {product: True for product in unique_products}

            # Toon checkboxen voor elk product
            selected_products = []
            cols = st.columns(min(3, len(unique_products)))  # Max 3 kolommen

            for i, product in enumerate(unique_products):
                col = cols[i % len(cols)]
                if col.checkbox(
                    product,
                    value=st.session_state.product_selections.get(product, True),
                    key=f"select_product_{i}"
                ):
                    st.session_state.product_selections[product] = True
                    selected_products.append(product)
                else:
                    st.session_state.product_selections[product] = False

            # Toon geselecteerde producten
            if selected_products:
                st.success(f"{len(selected_products)} producten geselecteerd: {', '.join(selected_products[:3])}{'...' if len(selected_products) > 3 else ''}")
            else:
                st.warning("Geen producten geselecteerd")

            # Genereer labels knop (alleen als er producten geselecteerd zijn)
            if selected_products and st.button("Genereer Verzendlabels", type="primary", use_container_width=True):
                with st.spinner("Verzendlabels worden gegenereerd..."):
                    try:
                        # Genereer labels met filters
                        labels = generate_shipping_labels(
                            df=df,
                            allowed_products=selected_products,
                            sort_order=sort_order,
                            start_date=datetime.combine(start_date, datetime.min.time()) if start_date else None,
                            end_date=datetime.combine(end_date, datetime.max.time()) if end_date else None
                        )

                        if not labels:
                            st.error("Geen geldige labels gevonden met de geselecteerde filters.")
                            return

                        st.success(f"{len(labels)} unieke verzendlabels gegenereerd!")

                        # Maak een tijdelijk bestand voor de PDF met datum in de naam
                        today = datetime.now().strftime("%Y-%m-%d")
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                            pdf_path = create_pdf_from_labels(labels, tmp_file.name)

                        # Lees de PDF voor download
                        with open(pdf_path, 'rb') as f:
                            pdf_bytes = f.read()

                        # Download knop met datum in filename
                        filename = f"verzendlabels_{today}.pdf"
                        st.download_button(
                            label="Download PDF",
                            data=pdf_bytes,
                            file_name=filename,
                            mime="application/pdf",
                            use_container_width=True
                        )

                        # Info over de PDF
                        st.info(f"PDF bevat {len(labels)} labels verdeeld over {(len(labels) + 23) // 24} pagina's (8×3 labels per pagina).")

                        # Voorbeeld van eerste paar labels
                        st.subheader("Voorbeeld van labels:")
                        for i, label in enumerate(labels[:5]):
                            st.text(label)
                        if len(labels) > 5:
                            st.text(f"... en nog {len(labels) - 5} labels")

                    except Exception as e:
                        st.error(f"Fout bij het genereren van labels: {e}")

    # Instructions
    st.markdown("---")
    st.markdown("### Instructies")
    st.markdown("""
    1. Upload een CSV-bestand met de volgende kolommen:
       - `company` (Bedrijfsnaam, optioneel)
       - `firstname` (Voornaam)
       - `lastname` (Achternaam)
       - `street` (Straatnaam)
       - `housenumber` (Huisnummer)
       - `housenumber_suffix` (Toevoeging, optioneel)
       - `zipcode` (Postcode)
       - `city` (Plaats)
       - `product` (Productnaam)
       - `paid_at` (Betaaldatum, voor sortering)

    2. Klik op "Genereer Verzendlabels" om de PDF te maken

    3. Download de gegenereerde PDF

    **Filtering:** Selecteer de producten die je wilt verwerken en stel datumbereik in indien nodig.
    """)

if __name__ == "__main__":
    main()
