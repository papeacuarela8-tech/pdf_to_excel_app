import os
import re
import sys
import statistics
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image

_ocr_reader = None

def get_ocr_reader():
    """Lazy loader for EasyOCR reader to avoid slow startup and memory bloat."""
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        _ocr_reader = easyocr.Reader(['es'], gpu=False)
    return _ocr_reader

def clean_currency(txt):
    if not txt:
        return ""
    # In Mexican invoices a comma is normally a thousands separator.  Replacing
    # every comma with a dot changed values such as $1,250.50 into 1.250.50.
    txt = str(txt).upper().replace('$', '').replace('S', '').replace(' ', '').strip()
    txt = re.sub(r'[^0-9,.-]', '', txt)
    if ',' in txt and '.' in txt:
        txt = txt.replace(',', '')
    elif ',' in txt:
        # Treat a final comma followed by one/two digits as a decimal separator;
        # otherwise it is a grouping separator.
        txt = txt.replace(',', '.') if re.search(r',\d{1,2}$', txt) else txt.replace(',', '')
    m = re.search(r'-?\d+(?:\.\d+)?', txt)
    return m.group(0) if m else ""


def _ocr_lines(entries):
    """Group OCR tokens into printed baselines without merging adjacent rows."""
    if not entries:
        return []

    # EasyOCR's boxes give us a useful, document-independent row tolerance.
    heights = [e['y1'] - e['y0'] for e in entries if e['y1'] > e['y0']]
    median_height = statistics.median(heights) if heights else 12
    tolerance = max(4, min(11, median_height * 0.45))
    lines = []
    for entry in sorted(entries, key=lambda e: e['yc']):
        if lines and abs(entry['yc'] - lines[-1]['yc']) <= tolerance:
            line = lines[-1]
            line['items'].append(entry)
            line['yc'] = sum(item['yc'] for item in line['items']) / len(line['items'])
        else:
            lines.append({'yc': entry['yc'], 'items': [entry]})
    for line in lines:
        line['items'].sort(key=lambda e: e['xc_rel'])
        line['text'] = ' '.join(e['text_clean'] for e in line['items'])
    return lines


def _money_value(value):
    try:
        return float(clean_currency(value))
    except (TypeError, ValueError):
        return 0.0


def _discount_fraction(value):
    """Return a discount as an Excel-ready fraction (5% -> 0.05)."""
    text = str(value or '').strip()
    numeric = _money_value(text)
    if '%' in text or numeric > 1:
        numeric /= 100
    return min(max(numeric, 0.0), 1.0)


def _safe_cell_text(value):
    """Prevent product text from being treated as a spreadsheet formula."""
    text = str(value or '')
    return f"'{text}" if text.startswith(('=', '+', '-', '@')) else text


def _extract_ocr_page_rows(entries, page_height, full_text, discount):
    """Read invoice rows from OCR geometry, preserving one printed row per product.

    Super Papelera CFDI invoices use a fixed column layout.  It is deliberately
    handled separately because its quantity/code columns are to the *left* of
    the description, unlike the generic SAT layouts handled below.
    """
    lines = _ocr_lines(entries)
    headers = [line for line in lines if "DESCRIP" in line['text'].upper()
               and any(word in line['text'].upper() for word in ("CANT", "PRECIO", "CODIGO", "CÓDIGO"))]
    header_y = headers[0]['yc'] if headers else page_height * .20
    footer_y = min((line['yc'] for line in lines if line['yc'] > header_y and any(
        marker in line['text'].upper() for marker in ("SUBTOTAL", "IMPORTE CON LETRA", "PAGAR", "QUEJAS"))), default=page_height * .90)
    super_papelera = "SUPER PAPELERA" in full_text.upper() or any(
        any(code in line['text'].upper() for code in ("CODIGO", "CÓDIGO"))
        and "UNID" in line['text'].upper() for line in lines)
    rows, previous = [], None
    for line in lines:
        if not header_y < line['yc'] < footer_y:
            continue
        cols = {key: [] for key in ('no', 'code', 'cve_sat', 'desc', 'qty', 'unit', 'price', 'discount', 'final', 'total')}
        for item in line['items']:
            x = item['xc_rel']
            if super_papelera:
                # No. | Cantidad | Unidad | Código | Cve. SAT | Descripción |
                # Precio Uni. | % Desc. | Prec. Final | Importe
                key = ('no' if x < .070 else 'qty' if x < .130 else 'unit' if x < .185 else
                       'code' if x < .275 else 'cve_sat' if x < .390 else 'desc' if x < .705 else
                       'price' if x < .790 else 'discount' if x < .850 else 'final' if x < .925 else 'total')
            else:
                key = ('no' if x < .08 else 'qty' if x < .14 else 'unit' if x < .19 else
                       'code' if x < .28 else 'desc' if x < .68 else 'price' if x < .80 else 'total')
            cols[key].append(item['text_clean'])
        desc = ' '.join(cols['desc']).strip()
        qty = _money_value(' '.join(cols['qty']))
        price = _money_value(' '.join(cols['price']))
        final_price = _money_value(' '.join(cols['final']))
        total = _money_value(' '.join(cols['total']))
        row_discount = _discount_fraction(' '.join(cols['discount'])) if cols['discount'] else discount
        if desc and (price > 0 or final_price > 0 or total > 0) and (qty > 0 or cols['code']):
            qty = qty or 1
            # Prefer the supplier's pre-discount unit price.  If it was not
            # recognised, reconstruct it from the final price and discount.
            price = price or (round(final_price / (1 - row_discount), 2) if final_price and row_discount < 1 else 0)
            price = price or round(total / qty, 2)
            final_price = final_price or round(price * (1 - row_discount), 2)
            total = total or round(final_price * qty, 2)
            previous = {"no": ''.join(cols['no']), "cantidad": qty, "unidad": ' '.join(cols['unit']) or "PZA",
                        "codigo": ''.join(cols['code']), "cve_sat": ''.join(cols['cve_sat']), "descripcion": desc,
                        "precio_unitario": price, "descuento": row_discount, "precio_final": final_price, "importe": total}
            rows.append(previous)
        elif previous and desc and not any(cols[key] for key in ('no', 'code', 'qty', 'unit', 'price', 'total')):
            previous['descripcion'] += ' ' + desc
    return rows

def extract_with_pdfplumber(pdf_path):
    """
    Extract line items from digital PDFs using pdfplumber native text & table extraction.
    Fast, highly accurate, and supports non-scanned invoices.
    """
    import pdfplumber
    
    all_prods = []
    invoice_number = "N/A"
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text() or ""
                full_text += text + "\n"
            
            # Find invoice / remision number in text
            inv_match = re.search(r'(?:REMISI[OÓ]N|Factura|Folio|Factura\s*No\.?|Folio\s*Fiscal)\s*:?\s*([A-Z0-9-]+)', full_text, re.IGNORECASE)
            if inv_match:
                invoice_number = inv_match.group(1).replace(" ", "")
            else:
                inv_match2 = re.search(r'\b([E|F|T]-\d+|\d{6,10})\b', full_text)
                if inv_match2:
                    invoice_number = inv_match2.group(1)

            # Check for global discount
            global_discount = 0.0
            disc_match = re.search(r'(\d+(?:\.\d+)?)\s*%\s*DESC', full_text, re.IGNORECASE)
            if disc_match:
                try:
                    global_discount = float(disc_match.group(1)) / 100.0
                except:
                    global_discount = 0.05

            # 1. Try structured table extraction
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    
                    header_row_idx = -1
                    col_map = {}
                    
                    for idx, row in enumerate(table[:6]):
                        row_str = " ".join([str(c or "").upper() for c in row if c])
                        if any(h in row_str for h in ["CANTIDAD", "CANT", "DESCRIP", "CONCEPTO", "PRECIO", "IMPORTE", "VALOR", "TOTAL"]):
                            header_row_idx = idx
                            for c_idx, cell in enumerate(row):
                                cell_txt = str(cell or "").upper().strip()
                                if "CANT" in cell_txt:
                                    col_map['cantidad'] = c_idx
                                elif "UNIDAD" in cell_txt or "UNID" in cell_txt or "UM" in cell_txt:
                                    col_map['unidad'] = c_idx
                                elif "CLAVE" in cell_txt or "CODIGO" in cell_txt or "CÓDIGO" in cell_txt or "NO. IDENT" in cell_txt:
                                    if 'codigo' not in col_map:
                                        col_map['codigo'] = c_idx
                                    else:
                                        col_map['cve_sat'] = c_idx
                                elif "DESCRIP" in cell_txt or "CONCEPTO" in cell_txt:
                                    col_map['descripcion'] = c_idx
                                elif "PRECIO" in cell_txt or "VALOR UNIT" in cell_txt or "P.U" in cell_txt:
                                    col_map['precio_unitario'] = c_idx
                                elif "DESC" in cell_txt:
                                    col_map['descuento'] = c_idx
                                elif "IMPORTE" in cell_txt or "TOTAL" in cell_txt:
                                    col_map['importe'] = c_idx
                            break
                    
                    if header_row_idx != -1 and 'cantidad' in col_map and 'descripcion' in col_map:
                        for row in table[header_row_idx + 1:]:
                            if not row or len(row) <= max(col_map.values()):
                                continue
                            
                            qty_val = clean_currency(row[col_map['cantidad']])
                            desc_val = str(row[col_map['descripcion']] or "").strip()
                            
                            if not qty_val or not re.match(r'^\d+(\.\d+)?$', qty_val) or not desc_val:
                                continue
                            
                            if any(term in desc_val.upper() for term in ["SUBTOTAL", "TOTAL", "IMPUESTO", "IVA", "TASA", "PAGO", "CADENA"]):
                                continue
                            
                            unit_val = str(row[col_map['unidad']] or "").strip() if 'unidad' in col_map else ""
                            code_val = str(row[col_map['codigo']] or "").strip() if 'codigo' in col_map else ""
                            sat_val = str(row[col_map['cve_sat']] or "").strip() if 'cve_sat' in col_map else ""
                            pu_val = clean_currency(row[col_map['precio_unitario']]) if 'precio_unitario' in col_map else ""
                            disc_val = _discount_fraction(row[col_map['descuento']]) if 'descuento' in col_map else global_discount
                            imp_val = clean_currency(row[col_map['importe']]) if 'importe' in col_map else ""
                            
                            all_prods.append({
                                "no": str(len(all_prods) + 1),
                                "cantidad": qty_val,
                                "unidad": unit_val,
                                "codigo": code_val,
                                "cve_sat": sat_val,
                                "descripcion": desc_val,
                                "precio_unitario": pu_val,
                                "descuento": disc_val,
                                "precio_final": pu_val,
                                "importe": imp_val
                            })

            # 2. If table extraction yielded no items, parse text line by line
            if not all_prods:
                for page in pdf.pages:
                    words = page.extract_words()
                    if not words:
                        continue
                    
                    lines_dict = {}
                    for w in words:
                        top_key = round(w['top'] / 4.0) * 4
                        lines_dict.setdefault(top_key, []).append(w)
                    
                    sorted_lines = sorted(lines_dict.items(), key=lambda x: x[0])
                    
                    for _, line_words in sorted_lines:
                        line_words.sort(key=lambda w: w['x0'])
                        line_text = " ".join([w['text'] for w in line_words])
                        
                        match = re.search(r'^(\d+(?:\.\d+)?)\s+([A-Z0-9]{1,10})?\s*(.*?)\s+\$?(\d+(?:\.\d+)?)\s*(?:\$?(\d+(?:\.\d+)?))?$', line_text)
                        if match:
                            qty, unit, desc, pu, imp = match.groups()
                            if not any(term in desc.upper() for term in ["SUBTOTAL", "TOTAL", "IMPUESTO", "IVA", "CADENA"]):
                                all_prods.append({
                                    "no": str(len(all_prods) + 1),
                                    "cantidad": qty,
                                    "unidad": unit or "",
                                    "codigo": "",
                                    "cve_sat": "",
                                    "descripcion": desc.strip(),
                                    "precio_unitario": pu,
                                    "descuento": str(global_discount),
                                    "precio_final": pu,
                                    "importe": imp or pu
                                })
    except Exception as e:
        print(f"pdfplumber extraction error: {e}")
        
    return all_prods, invoice_number

def extract_with_ocr(pdf_path, temp_dir):
    """
    Extract line items from scanned PDF images using pypdfium2 page rendering + EasyOCR.
    Supports Super Papelera and SAT Mexican invoice layouts dynamically.
    """
    import pypdfium2 as pdfium
    ocr_reader = get_ocr_reader()
    
    all_prods = []
    invoice_number = "N/A"
    
    pdf = pdfium.PdfDocument(pdf_path)
    img_paths = []
    
    try:
        for page_idx, page in enumerate(pdf):
            img = page.render(scale=2).to_pil()
            img_path = os.path.join(temp_dir, f"ocr_page_{page_idx+1}_{os.getpid()}.jpg")
            img.save(img_path, "JPEG")
            img_paths.append(img_path)

        for page_num, img_path in enumerate(img_paths, start=1):
            with Image.open(img_path) as pil_img:
                width, height = pil_img.size

            results = ocr_reader.readtext(img_path)
            
            entries = []
            full_ocr_text = ""
            for bbox, text, conf in results:
                full_ocr_text += text + " "
                yc = sum(pt[1] for pt in bbox) / 4.0
                xc = sum(pt[0] for pt in bbox) / 4.0
                
                entries.append({
                    "xc_rel": xc / float(width),
                    "yc": yc,
                    "y0": min(pt[1] for pt in bbox),
                    "y1": max(pt[1] for pt in bbox),
                    "text_clean": text.strip(),
                    "text": text,
                    "conf": conf
                })

            # Search for invoice / remision / folio number
            inv_match = re.search(r'(?:REMISI[OÓ]N|FACTURA|FOLIO)\s*:?\s*([A-Z0-9-]+)', full_ocr_text, re.IGNORECASE)
            if inv_match:
                invoice_number = inv_match.group(1).replace(" ", "")
            else:
                inv_match2 = re.search(r'\b([E|F|T]-\d+|\d{6,10})\b', full_ocr_text)
                if inv_match2:
                    invoice_number = inv_match2.group(1)

            # Check for global discount
            global_discount = 0.0
            disc_match = re.search(r'(\d+(?:\.\d+)?)\s*%\s*DESC', full_ocr_text, re.IGNORECASE)
            if disc_match:
                try:
                    global_discount = float(disc_match.group(1)) / 100.0
                except:
                    global_discount = 0.05

            # Use geometric table parsing first.  The legacy parser below is
            # retained for reference, but its 20px grouping merged neighbours.
            geometric_rows = _extract_ocr_page_rows(entries, height, full_ocr_text, global_discount)
            if geometric_rows:
                all_prods.extend(geometric_rows)
                continue

            # Find table top header Y
            header_y = 0
            for e in entries:
                if any(h in e["text_clean"].upper() for h in ["CODIGO", "CÓDIGO", "DESCRIPCION", "DESCRIPCIÓN", "CANT.", "PRECIO", "TOTAL"]):
                    header_y = e["yc"]
                    break
            
            table_start_y = header_y + 15 if header_y > 0 else height * 0.20
            
            # Find footer Y
            footer_y = height * 0.88
            for e in entries:
                if e["yc"] > table_start_y and any(f in e["text_clean"].upper() for f in ["SUBTOTAL", "IMPORTE CON LETRA", "PAGARÉ", "QUEJAS"]):
                    if e["yc"] < footer_y:
                        footer_y = e["yc"]
            
            table_entries = [e for e in entries if table_start_y <= e["yc"] < footer_y]
            
            # Group into horizontal rows (vertical tolerance 20px)
            row_groups = []
            for e in sorted(table_entries, key=lambda x: x["yc"]):
                placed = False
                for group in row_groups:
                    avg_y = sum(item["yc"] for item in group) / len(group)
                    if abs(e["yc"] - avg_y) <= 20:
                        group.append(e)
                        placed = True
                        break
                if not placed:
                    row_groups.append([e])

            # Determine layout type (Super Papelera vs SAT/Generic)
            is_super_papelera = False
            for group in row_groups:
                for e in group:
                    if 0.57 <= e["xc_rel"] <= 0.72 and any(u in e["text_clean"].upper() for u in ["PZA", "PAQ", "CJA", "KG", "MTR"]):
                        is_super_papelera = True
                        break

            for group in row_groups:
                group.sort(key=lambda x: x["xc_rel"])
                
                no_str = ""
                code_str = ""
                desc_parts = []
                qty_str = ""
                unit_str = ""
                price_str = ""
                total_str = ""

                if is_super_papelera:
                    for e in group:
                        xr = e["xc_rel"]
                        txt = e["text_clean"]
                        
                        if xr < 0.07:
                            if txt.isdigit():
                                no_str = txt
                        elif 0.07 <= xr < 0.19:
                            if txt.isdigit() or len(txt) >= 4:
                                code_str = txt
                        elif 0.19 <= xr < 0.57:
                            if not any(term in txt.upper() for term in ["CODIGO", "DESCRIPCION", "CANT", "UNID", "PRECIO", "TOTAL"]):
                                desc_parts.append(txt)
                        elif 0.57 <= xr < 0.63:
                            clean_q = re.sub(r'[^\d.]', '', txt)
                            if clean_q:
                                qty_str = clean_q
                        elif 0.63 <= xr < 0.71:
                            unit_match = re.search(r'([A-Z]{2,4})', txt)
                            if unit_match:
                                unit_str = unit_match.group(1)
                            else:
                                unit_str = txt
                        elif 0.71 <= xr < 0.82:
                            price_str = txt
                        elif 0.82 <= xr < 0.98:
                            total_str = txt
                else:
                    for e in group:
                        xr = e["xc_rel"]
                        txt = e["text_clean"]

                        if xr < 0.08:
                            if txt.isdigit():
                                no_str = txt
                        elif 0.08 <= xr < 0.14:
                            clean_q = re.sub(r'[^\d.]', '', txt)
                            if clean_q:
                                qty_str = clean_q
                        elif 0.14 <= xr < 0.19:
                            if txt.upper() not in ["IMPUESTO", "IVA", "002"]:
                                unit_str = txt
                        elif 0.19 <= xr < 0.28:
                            if txt.isdigit() or len(txt) > 3:
                                code_str = txt
                        elif 0.28 <= xr < 0.68:
                            if not any(term in txt.upper() for term in ["IMPUESTO", "TIPO FACTOR", "TASA", "IMPORTE", "SUBTOTAL", "TOTAL"]):
                                desc_parts.append(txt)
                        elif 0.68 <= xr < 0.80:
                            if '$' in txt or re.search(r'\d+\.\d+', txt):
                                price_str = txt
                        elif 0.80 <= xr < 0.98:
                            if '$' in txt or re.search(r'\d+\.\d+', txt):
                                total_str = txt

                desc = " ".join(desc_parts).strip()
                if not desc:
                    continue

                def clean_val(v):
                    if not v:
                        return 0.0
                    v = v.replace('$', '').replace('S', '').replace('s', '').replace(',', '.').strip()
                    v = re.sub(r'[^\d.]', '', v)
                    try:
                        return float(v)
                    except:
                        return 0.0

                p_val = clean_val(price_str)
                t_val = clean_val(total_str)

                try:
                    q_val = float(qty_str)
                    if q_val.is_integer():
                        q_val = int(q_val)
                except:
                    q_val = 1

                if desc and (p_val > 0 or t_val > 0):
                    if p_val == 0.0 and t_val > 0 and q_val > 0:
                        p_val = round(t_val / float(q_val), 2)
                    elif t_val == 0.0 and p_val > 0 and q_val > 0:
                        t_val = round(p_val * float(q_val), 2)

                    all_prods.append({
                        "no": no_str,
                        "cantidad": q_val,
                        "unidad": unit_str or "PZA",
                        "codigo": code_str,
                        "cve_sat": "",
                        "descripcion": desc,
                        "precio_unitario": p_val,
                        "descuento": global_discount,
                        "precio_final": p_val,
                        "importe": t_val
                    })
    finally:
        for img_path in img_paths:
            try:
                if os.path.exists(img_path):
                    os.remove(img_path)
            except:
                pass
                
    return all_prods, invoice_number

def parse_pdf_invoice(pdf_path, output_excel_path, markup_percent=65):
    """
    Main function to parse PDF invoices (digital or scanned) and output
    a styled Excel spreadsheet matching Papelería Acuarela design.
    """
    temp_dir = os.path.dirname(output_excel_path) or os.getcwd()
    try:
        markup_percent = float(markup_percent)
    except (TypeError, ValueError):
        markup_percent = 65.0
    markup_multiplier = 1 + markup_percent / 100
    
    # 1. First try fast native text extraction with pdfplumber
    all_prods, invoice_number = extract_with_pdfplumber(pdf_path)
    
    # 2. Fallback to full-page OCR using pypdfium2 + EasyOCR if no products were found
    if not all_prods:
        all_prods, invoice_ocr = extract_with_ocr(pdf_path, temp_dir)
        if invoice_ocr and invoice_ocr != "N/A":
            invoice_number = invoice_ocr
            
    if not all_prods:
        raise ValueError("No se pudieron extraer los productos de la factura PDF. Por favor verifica que el archivo sea una factura legible.")

    # 3. Assign sequential numbers
    # Supplier line numbers restart on each page.  Keep the actual document
    # order, instead of sorting every page's “1” together with every other 1.
    for current_no, p in enumerate(all_prods, start=1):
        p["assigned_no"] = current_no

    # 4. Use the user's workbook as the actual output template.  This keeps
    # every visual detail (fonts, colours, borders, widths and row heights)
    # instead of attempting to reproduce it in code.
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "formato_de_tabla_excel.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError("No se encontró la plantilla formato_de_tabla_excel.xlsx.")

    wb = openpyxl.load_workbook(template_path)
    sheet = wb["Precios de Venta"] if "Precios de Venta" in wb.sheetnames else wb.active
    data_start_row = 5

    # Capture the product-row style before removing the template's example data.
    row_style = [copy(sheet.cell(data_start_row, column)._style) for column in range(1, 9)]
    row_height = sheet.row_dimensions[data_start_row].height
    existing_data_rows = max(sheet.max_row - data_start_row + 1, 0)
    if existing_data_rows:
        sheet.delete_rows(data_start_row, existing_data_rows)

    sheet["A2"] = f"Factura {invoice_number} | Papelería Acuarela | Cálculo de Precio de Venta (Costo + {markup_percent:g}%)"
    sheet["H4"] = f"Precio Venta (Costo + {markup_percent:g}%)"

    for idx, p in enumerate(all_prods, start=data_start_row):
        for col_idx in range(1, 9):
            sheet.cell(row=idx, column=col_idx)._style = copy(row_style[col_idx - 1])
        sheet.row_dimensions[idx].height = row_height
        try:
            qty = float(p["cantidad"])
            if qty.is_integer():
                qty = int(qty)
        except:
            qty = 1

        try:
            pu = float(p["precio_unitario"])
        except:
            pu = 0.0

        try:
            disc = _discount_fraction(p["descuento"])
        except:
            disc = 0.0

        sheet.cell(row=idx, column=1, value=qty)
        sheet.cell(row=idx, column=2, value=_safe_cell_text(p.get("unidad", "")))
        sheet.cell(row=idx, column=3, value=_safe_cell_text(p.get("descripcion", "")))
        sheet.cell(row=idx, column=4, value=pu)
        sheet.cell(row=idx, column=5, value=disc)
        
        sheet.cell(row=idx, column=6, value=f"=D{idx}*(1-E{idx})")
        sheet.cell(row=idx, column=7, value=f"=A{idx}*F{idx}")
        sheet.cell(row=idx, column=8, value=f"=F{idx}*{markup_multiplier}")

    wb.save(output_excel_path)
    return invoice_number
