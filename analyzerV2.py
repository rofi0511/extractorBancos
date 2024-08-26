import re
import pandas as pd
import pdfplumber
from tkinter import Tk, Button, Label, filedialog, messagebox, StringVar, OptionMenu

def extract_pdf_text(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text()
    return full_text

def extract_saldo_inicial(full_text):
    saldo_pattern = re.compile(r'S\s*aldo\s*inicial.*?\$\d{1,3}(?:,\d{3})*\.\d{2}', re.IGNORECASE)
    match = saldo_pattern.search(full_text)
    if match:
        saldo_inicial = float(re.findall(r'\$\d{1,3}(?:,\d{3})*\.\d{2}', match.group())[0].replace('$', '').replace(',', ''))
        return saldo_inicial
    return None

def extract_amounts_adjusted(description):
    amounts = re.findall(r'\d{1,3}(?:,\d{3})*\.\d{2}', description)
    if len(amounts) >= 2:
        deposito_retiro = float(amounts[-2].replace(',',''))
        saldo = float(amounts[-1].replace(',',''))
        return deposito_retiro, saldo
    elif len(amounts) == 1:
        return None, float(amounts[0].replace(',',''))
    else:
        return None, None

def refine_and_capture_movements(full_text):
    lines = full_text.split('\n')
    movements = []
    date_pattern = re.compile(r'^\d{2} \w{3}')
    amount_pattern = re.compile(r'\$\d{1,3}(?:,\d{3})*\.\d{2}')
    combined_line = ""
    skip_next_line = False

    for i, line in enumerate(lines):
        line = line.strip()

        if "Saldo" in line or "final" in line or "Comisionescobradas" in line or not line:
            continue

        if skip_next_line:
            skip_next_line = False
            continue

        if "DEPOSITOS SBC CAMARA" in line:
            combined_line = line.strip() + " " + lines[i + 1].strip()
            skip_next_line = True
            movements.append(combined_line)
            combined_line = ""
        elif date_pattern.match(line[:6]):
            if combined_line:
                movements.append(combined_line.strip())
                combined_line = ""
            combined_line = line.strip()
        elif amount_pattern.search(line):
            combined_line += " " + line.strip()
        elif combined_line:
            movements.append(combined_line.strip())
            combined_line = ""
        else:
            continue

    if combined_line:
        movements.append(combined_line.strip())
    
    return movements

def classify_movements_with_saldo_initial(movements, saldo_inicial):
    movements_data = []
    previous_saldo = saldo_inicial

    for movement in movements:
        parts = movement.split()
        if len(parts) > 4:
            try:
                oper_date = " ".join(parts[:2])
                concepto = " ".join(parts[2:-3]).strip()
                retiro_deposito = float(parts[-2].replace('$','').replace(',',''))
                saldo = float(parts[-1].replace('$','').replace(',',''))
            

                if saldo < previous_saldo:
                    retiro = retiro_deposito
                    deposito = 0
                else:
                    retiro = 0
                    deposito = retiro_deposito

                movement_info = {
                    "Fecha": oper_date,
                    "Concepto": concepto,
                    "Depósito": deposito if deposito != 0 else 0,
                    "Retiro": retiro if retiro != 0 else 0,
                    "Saldo": saldo
                }

                movements_data.append(movement_info)
                previous_saldo = saldo
            except ValueError:
                continue
    return pd.DataFrame(movements_data)


def extract_date_info(full_text):
    date_info_pattern = re.compile(r'del \d{2} al \d{2} de (\w+) (\d{4})')
    match = date_info_pattern.search(full_text)
    if match:
        month = match.group(1)
        year = match.group(2)
        return month, year
    return None, None

def clean_text(full_text):
    lines = full_text.split('\n')
    cleaned_lines = []

    for line in lines:
        if not any(keyword in line for keyword in [
            "000180.B07CHDA008.OD.0731.01",
            "ESTADOS DE CUENTA AL",
            "CLIENTE:",
            "Página:",
            "GRUPO SUNEGO DE PUEBLA SA DE CV",
            "DETALLE DE OPERACIONES",
            "FECHA CONCEPTO RETIROS DEPOSITOS SALDO"
        ]):
            cleaned_lines.append(line)
    return '\n'.join(cleaned_lines)

def extract_movements_section(full_text):
    start_marker = "Detalle de movimientos cuenta de cheques."
    end_marker = "Detalles de movimientos Dinero Creciente Santander."
    start_index = full_text.find(start_marker)
    end_index = full_text.find(end_marker)
    if start_index != -1 and end_index != -1:
        return full_text[start_index:end_index]
    else:
        return ""
    
def extract_and_format_movements_flexible(movements_section):
    saldo_pattern = re.compile(r'SALDOFINALDELPERIODOANTERIOR:\s*\$(\d{1,3}(?:,\d{3})*\.\d{2})')
    match = saldo_pattern.search(movements_section)
    saldo_anterior = float(match.group(1).replace(',','')) if match else None

    movement_pattern = re.compile(r'(\d{2}-[A-Z]{3}-\d{4})\s+(\d+)([A-Z\s]+)\s+(.+?)\s+(\d{1,3}(?:,\d{3})*\.\d{2})\s+(\d{1,3}(?:,\d{3})*\.\d{2})')
    movimientos = []

    for m in movement_pattern.finditer(movements_section):
        fecha = m.group(1)
        descripcion = m.group(4).strip()
        monto = float(m.group(5).replace(',',''))
        saldo = float(m.group(6).replace(',',''))

        movimientos.append({
            'Fecha': fecha,
            'Descripción': descripcion,
            'Monto': monto,
            'Saldo': saldo
        })

    return saldo_anterior, pd.DataFrame(movimientos)

def classify_movements(df_movimientos, saldo_anterior):
    df_movimientos['Tipo'] = ''
    df_movimientos['Retiro'] = 0.0
    df_movimientos['Depósito'] = 0.0

    for i, row in df_movimientos.iterrows():
        if row['Saldo'] < saldo_anterior:
            df_movimientos.at[i, 'Tipo'] = 'Retiro'
            df_movimientos.at[i, 'Retiro'] = row['Monto']
        else:
            df_movimientos.at[i, 'Tipo'] = 'Depósito'
            df_movimientos.at[i, 'Depósito'] = row['Monto']

        saldo_anterior = row['Saldo']
    
    return df_movimientos

def extract_movements_azteca(full_text):
    start_marker = "Detalle de movimientos realizados"
    end_marker = "Revise cuidadosamente éste Estado de Cuenta."
    start_index = full_text.find(start_marker)
    end_index = full_text.find(end_marker)

    if start_index != -1 and end_index != -1:
        return full_text[start_index:end_index]
    else:
        return ""
    
def extract_movements_inbursa(full_text):
    start_marker = "Detalle de movimientos"
    end_marker = "Si desea recibir pagos a través"
    start_index = full_text.find(start_marker)
    end_index = full_text.find(end_marker)

    if start_index != -1 and end_index != -1:
        return full_text[start_index:end_index]
    else:
        return ""
    
def process_banamex_pdf(cleaned_text):
    lines = cleaned_text.split('\n')
    movements = []
    current_movement = {
        "Fecha": "",
        "Concepto": "",
        "Retiro": 0.0,
        "Depósito": 0.0,
        "Saldo": 0.0
    }
    combine_next_line = False

    for line in lines:
        line = line.strip()

        if "SALDO MINIMO REQUERIDO" in line:
            if current_movement["Fecha"]:
                movements.append(current_movement)
            break

        if re.match(r'\d{2} \w{3}', line[:6]) and not combine_next_line:
            if current_movement["Fecha"]:
                movements.append(current_movement)
            current_movement = {
                "Fecha": "",
                "Concepto": "",
                "Retiro": 0.0,
                "Depósito": 0.0,
                "Saldo": 0.0
            }
            current_movement["Fecha"] = line[:6]
            current_movement["Concepto"] = line[7:].strip()
        else:
            if re.search(r'\d{1,3}(?:,\d{3})*\.\d{2}', line):
                numbers = re.findall(r'\d{1,3}(?:,\d{3})*\.\d{2}', line)
                if len(numbers) == 3:
                    current_movement["Retiro"] = float(numbers[-3].replace(',', ''))
                    current_movement["Saldo"] = float(numbers[-1].replace(',', ''))
                    current_movement["Depósito"] = 0.0
                elif len(numbers) == 2:
                    current_movement["Saldo"] = float(numbers[-1].replace(',', ''))
                    monto = float(numbers[-2].replace(',', ''))
                elif len(numbers) == 1:
                    monto = float(numbers[-1].replace(',', ''))
                    current_movement["Saldo"] = 0.0

                concepto = current_movement["Concepto"].upper()
                if any(keyword in concepto for keyword in ["PAGO RECIBIDO", "ABONO", "DEPOSITO", "TRASPASO REF"]):
                    current_movement["Depósito"] = monto
                elif any(keyword in concepto for keyword in [
                    "PAGO A", "COMPRA", "RETIRO", "COMISION", "IVA COMISION", 
                    "DOMI AMERICAN EXPRESS", "COBRO IMP TPV GPRS", "COBRO COMI TPV GPRS", 
                    "COMPRA INVERSION INTEGRAL", "PAGO INTERBANCARIO A BBVA MEXICO", 
                    "COBRO IMP COM CUOT BJA FAC", "COBRO COM CUOT BJA FAC",
                    "PAGO INTERBANCARIO A BANORTE", "PAGO INTERBANCARIO A SANTANDER",
                    "PAGO INTERBANCARIO A BAJIO"]):
                    current_movement["Retiro"] = monto
                else:
                    current_movement["Retiro"] = 0.0
                    current_movement["Depósito"] = 0.0
                combine_next_line = False
            else:
                current_movement["Concepto"] += " " + line.strip()
                combine_next_line = True

    if current_movement["Fecha"] and current_movement not in movements:
        movements.append(current_movement)

    df = pd.DataFrame(movements)
    return df

def process_bancoazte_pdf(full_text):
    lines = full_text.split('\n')
    movements_data = []

    for line in lines:
        if re.match(r'\d{4}-\d{2}-\d{2}', line):
            parts = line.split()
            if len(parts) > 5:
                fecha_op = parts[0]
                concepto = " ".join(parts[4:-3])
                cargo = float(parts[-3].replace(',','')) if float(parts[-3].replace(',','')) != 0.00 else 0.0
                abono = float(parts[-2].replace(',','')) if float(parts[-2].replace(',','')) != 0.00 else 0.0

                movements_info = {
                    "Fecha Operación": fecha_op,
                    "Concepto": concepto,
                    "Cargo": cargo,
                    "Abono": abono
                }
                movements_data.append(movements_info)
            
    return pd.DataFrame(movements_data)

def process_bancomer_pdf(full_text):
    lines = full_text.split('\n')

    movements = []
    movements_data = []

    date_pattern = re.compile(r'\d{2}/\w{3}')

    for line in lines:
        if date_pattern.match(line.strip()):
            movements.append(line.strip())

    for movement in movements:
        parts = movement.split()
        if len(parts) > 2:
            oper_date = parts[0]
            amounts = [part for part in parts if re.match(r'\d{1,3}(?:,\d{3})*\.\d{2}', part)]
            description_end_index = parts.index(amounts[0]) if amounts else len(parts)
            description = " ".join(parts[2:description_end_index])

            cargo = "0"
            abono = "0"
            
            if any(keyword in description.lower() for keyword in ["abono", "depósito", "traspaso", "recibidos"]):
                abono = amounts[0]
            else:
                cargo = amounts[0]

            movement_info = {
                "Operación": oper_date,
                "Descripción": description,
                "Cargos": cargo,
                "Abonos": abono
            }

            movements_data.append(movement_info)

    return pd.DataFrame(movements_data)

def process_banorte_pdf(full_text):
    lines = full_text.split('\n')

    date_pattern = re.compile(r'^\d{2}-[A-Z]{3}-\d{2}')
    amount_pattern = re.compile(r'\d{1,3}(?:,\d{3})*\.\d{2}')

    movements = []
    current_entry = None

    for line in lines:
        if date_pattern.match(line):
            if current_entry:
                movements.append(current_entry)
            current_entry = {"Fecha": line[:9], "Descripción": line[9:].strip()}
        elif current_entry and amount_pattern.search(line):
            amounts = amount_pattern.findall(line)
            if len(amounts) == 2:
                current_entry["Monto"] = float(amounts[0].replace(',', ''))
                current_entry["Saldo"] = float(amounts[1].replace(',', ''))
            elif len(amounts) == 1:
                if "Monto" not in current_entry:
                    current_entry["Monto"] = float(amounts[0].replace(',', ''))
                else:
                    current_entry["Saldo"] = float(amounts[0].replace(',', ''))
        elif current_entry:
            current_entry["Descripción"] += " " + line.strip()
    
    if current_entry:
        movements.append(current_entry)

    df_movements = pd.DataFrame(movements)

    if df_movements.empty:
        print("No se encontraron movimientos.")
        return pd.DataFrame()

    saldo_anterior_entry = df_movements[df_movements['Descripción'].str.contains('SALDO ANTERIOR', case=False)]
    if not saldo_anterior_entry.empty:
        saldo_inicial = float(re.findall(r'\d{1,3}(?:,\d{3})*\.\d{2}', saldo_anterior_entry['Descripción'].values[0])[0].replace(',', ''))
    else:
        saldo_inicial = None
    df_movements_filtered = df_movements[~df_movements['Descripción'].str.contains('SALDO ANTERIOR', case=False)]

    df_movements_filtered[['Depósito/Retiro', 'Saldo']] = df_movements_filtered['Descripción'].apply(lambda x: pd.Series(extract_amounts_adjusted(x)))

    saldo_actual = saldo_inicial  

    df_movements_filtered['Tipo'] = None
    for i, row in df_movements_filtered.iterrows():
        if row['Saldo'] < saldo_actual:
            df_movements_filtered.at[i, 'Tipo'] = 'Retiro'
        else:
            df_movements_filtered.at[i, 'Tipo'] = 'Depósito'
        saldo_actual = row['Saldo']

    df_movements_filtered['Retiro'] = df_movements_filtered.apply(lambda x: x['Depósito/Retiro'] if x['Tipo'] == 'Retiro' else 0, axis=1)
    df_movements_filtered['Depósito'] = df_movements_filtered.apply(lambda x: x['Depósito/Retiro'] if x['Tipo'] == 'Depósito' else 0, axis=1)

    df_movements_final = df_movements_filtered.drop(columns=['Depósito/Retiro', 'Monto', 'Saldo', 'Tipo'])

    return df_movements_final

def process_banregio_pdf(full_text):
    month, year = extract_date_info(full_text)
    
    if not month or not year:
        messagebox.showwarning("Advertencia", "No se pudo encontrar la información del mes y año en el estado de cuenta.")
        return pd.DataFrame()
    
    lines = full_text.split('\n')
    movements = []
    movements_data = []

    date_pattern = re.compile(r'^\d{2}$')
    amount_pattern = re.compile(r'\d{1,3}(?:,\d{3})*\.\d{2}')

    for line in lines:
        parts = line.split()
        if len(parts) > 3 and date_pattern.match(parts[0]):
            movements.append(line.strip())

    for movement in movements:
        parts = movement.split()

        if len(parts) > 3:
            day = parts[0]
            cargos = "0"
            abonos = "0"
            description_parts = []

            for i, part in enumerate(parts[1:], start=1):
                if amount_pattern.match(part):
                    if "TRA" in movement and cargos == "0":  
                        cargos = part
                    elif "INT" in movement and abonos == "0": 
                        abonos = part
                    break
                else:
                    description_parts.append(part)

            description = " ".join(description_parts).strip()

            full_date = f"{day}/{month}/{year}"

            movement_info = {
                "Fecha": full_date,
                "Descripción": description,
                "Cargos": cargos,
                "Abonos": abonos
            }

            movements_data.append(movement_info)    

    return pd.DataFrame(movements_data)

def process_santander_pdf(full_text):
    movements_section = extract_movements_section(full_text)
    saldo_anterior, df_movimientos = extract_and_format_movements_flexible(movements_section)
    if saldo_anterior is None or df_movimientos.empty:
        messagebox.showwarning("Advertencia", "No se pudo extraer el saldo inicial o no se encontraron movimientos.")
        return pd.DataFrame()

    df_classified = classify_movements(df_movimientos, saldo_anterior)

    df_summary = df_classified[['Fecha', 'Descripción', 'Retiro', 'Depósito']].copy()

    return df_summary

def process_inbursa_pdf(full_text):
    lines = full_text.split('\n')
    movements_data = []
    balance_inicial = None
    saldo_anterior = None

    for line in lines:
        line = line.strip()

        if not line:
            continue

        if 'BALANCE INICIAL' in line:
            balance_inicial = float(re.findall(r'\d{1,3}(?:,\d{3})*\.\d{2}', line)[0].replace(',',''))
            saldo_anterior = balance_inicial
            continue

        if re.match(r'^[A-Z]{3} \d{2}', line):
            parts = line.split()

            if len(parts) < 5:
                print(f"Línea ignorada por tener menos de 5 partes: {line}")
                continue

            try:
                fecha = " ".join(parts[:2])
                concepto = " ".join(parts[2:-2])
                monto = float(parts[-2].replace(',', ''))
                saldo = float(parts[-1].replace(',', ''))

                if saldo > saldo_anterior:
                    abonos = monto
                    cargos = 0.0

                else:
                    cargos = monto
                    abonos = 0.0

                movement_info = {
                    "Fecha": fecha,
                    "Concepto": concepto,
                    "Cargos": cargos,
                    "Abonos": abonos
                }
                movements_data.append(movement_info)

                saldo_anterior = saldo
            except ValueError as e:
                print(f"Error al procesar la linea: {line}, Error: {e}")
                continue

    return pd.DataFrame(movements_data)

def process_scotiabank_pdf(full_text):
    saldo_inicial = extract_saldo_inicial(full_text)
    if saldo_inicial is None:
        messagebox.showwarning("Advertencia", "No se pudo extraer el saldo inicial.")
        return pd.DataFrame()
    
    refined_movements = refine_and_capture_movements(full_text)
    df_classified = classify_movements_with_saldo_initial(refined_movements, saldo_inicial)
    
    messagebox.showwarning(
        "Revisión Necesaria",
        "Revisar el archivo Excel al final, ya que movimientos que estén muy abajo del archivo PDF no se registraran bien."
    )

    return df_classified

def process_pdf():
    pdf_path = filedialog.askopenfilename(title="Selecciona el archivo PDF", filetypes=[("PDF files", "*.pdf")])
    
    if not pdf_path:
        messagebox.showwarning("Advertencia", "No seleccionaste ningún archivo PDF.")
        return
    
    full_text = extract_pdf_text(pdf_path)
    cleaned_text = clean_text(full_text)

    if selected_bank.get() == "BANAMEX":
        df_movements = process_banamex_pdf(cleaned_text)
    elif selected_bank.get() == "BANCOAZTE":
        df_movements = process_bancoazte_pdf(full_text)
    elif selected_bank.get() == "BANCOMER":
        df_movements = process_bancomer_pdf(full_text)
    elif selected_bank.get() == "BANORTE":
        df_movements = process_banorte_pdf(full_text)
    elif selected_bank.get() == "BANREGIO":
        df_movements = process_banregio_pdf(full_text)
    elif selected_bank.get() == "INBURSA":
        df_movements = process_inbursa_pdf(full_text)
    elif selected_bank.get() == "SANTANDER":
        df_movements = process_santander_pdf(full_text)
    elif selected_bank.get() == "SCOTIABANK":
        df_movements = process_scotiabank_pdf(full_text)
    else:
        messagebox.showwarning("Advertencia", f"El análisis para {selected_bank.get()} no está implementado.")
        return
    
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Guardar archivo como")
    
    if save_path:
        df_movements.to_excel(save_path, index=False)
        messagebox.showinfo("Éxito", f"Archivo exportado a: {save_path}")

root = Tk()
root.title("Extractor de Movimientos Financieros")
root.geometry("300x200")

welcome_label = Label(root, text="Bienvenido a la aplicación", font=("Helvetica", 14))
welcome_label.pack(pady=10)

selected_bank = StringVar(root)
selected_bank.set("BANAMEX")

bank_menu = OptionMenu(root, selected_bank, "BANAMEX","BANCOAZTE", "BANCOMER", "BANORTE", "BANREGIO", "INBURSA", "SANTANDER", "SCOTIABANK")
bank_menu.pack(pady=10)

process_button = Button(root, text="Seleccionar y procesar PDF", command=process_pdf, font=("Helvetica", 12))
process_button.pack(pady=20)

root.mainloop()
