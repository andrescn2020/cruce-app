from flask import Flask, render_template, request, redirect, send_from_directory
import os
import PyPDF2
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side

app = Flask(__name__)


# Ruta para la carga de archivos y procesamiento
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["pdf_file"]
        if file:
            # Verificar si la carpeta 'uploads' existe, y si no, crearla
            if not os.path.exists("uploads"):
                os.makedirs("uploads")

            file_path = os.path.join("uploads", file.filename)
            file.save(file_path)
            try:
                # Aquí comienza tu código de procesamiento del PDF
                with open(file_path, "rb") as pdf_file:
                    reader = PyPDF2.PdfReader(pdf_file)
                    texto = "".join(page.extract_text() + "\n" for page in reader.pages)
                    texto = texto.splitlines()

                capturar = False
                movimientos = []
                movimiento = {}
                banco = texto[1]

                for linea in texto:
                    # print(linea)
                    if "F.de Pago" in linea:
                        # print(linea)
                        capturar = False
                        # Extraer fecha y nro de liquidación aca
                        # Extraer número CBU (6 o 7 dígitos)
                        cbu_match = re.search(
                            r"\d{1,3}(\.\d{3})*,\d+\-?\s+(\d+)", linea
                        )
                        # print(cbu_match)
                        nro_cbu = (
                            cbu_match.group(1)
                            if cbu_match
                            else "No se encontró Número de Liquidación"
                        )

                        # Extraer fecha de liquidación (DD/MM/AAAA)
                        fecha_match = (
                            re.search(
                                r"(\d{2}/\d{2}/\d{4})", linea.split("Nro. Liq:")[1]
                            )
                            if "Liq:" in linea
                            else None
                        )
                        if cbu_match:
                            movimiento["Liquidacion"] = cbu_match.group(2) + ".00"
                        else:
                            print("No se encontró ninguna coincidencia.")
                        fecha_liq = (
                            fecha_match.group(1)
                            if fecha_match
                            else "No se encontró fecha"
                        )
                        # Captura la fecha después de "Nro. Liq:"
                        movimiento["Fecha"] = fecha_liq
                        if nro_cbu == "No se encontró Número de Liquidación":

                            pass
                        else:
                            movimiento["Liquidacion"] = round(
                                float(movimiento["Liquidacion"])
                            )
                            movimientos.append(movimiento.copy())
                            # print(movimiento["Liquidacion"])
                            movimiento = {}

                    if "VENTAS" in linea or "QR" in linea or "AJUSTE" in linea:
                        capturar = True

                    if capturar:
                        partes = linea.split("$")
                        if len(partes) < 2:
                            pass
                        else:
                            # print(partes)
                            if "-" in partes[1]:
                                valor = partes[1].strip().replace("Fecha", "")
                                valor = valor.replace("-", "")
                                valor = valor.replace(".", "").replace(",", ".")
                                if "/" in valor:
                                    pass
                                else:
                                    valor = round(float(valor), 2) * -1
                                    partes[1] = valor
                            elif "Fecha" in partes[1]:
                                valor = partes[1].strip().replace("Fecha", "")
                                valor = valor.replace(".", "").replace(",", ".")
                                valor = round(float(valor), 2)
                                partes[1] = valor
                            else:
                                if "F.de Pago" in linea:
                                    pass
                                else:
                                    valor = (
                                        partes[1]
                                        .replace(".", "")
                                        .replace(",", ".")
                                        .replace("Fecha", "")
                                    )
                                    valor = round(float(valor), 2)
                                    partes[1] = valor
                            movimiento[partes[0]] = partes[1]
                            # print(movimiento)

                # print(movimientos)
                # Crear un DataFrame con los datos
                df_total = pd.DataFrame(movimientos)

                # Reemplazar valores NaN por 0
                df_total = df_total.fillna(0)

                # Identificar las columnas que comienzan con "QR"
                columnas_qr = [col for col in df_total.columns if col.startswith("QR")]

                # Filtrar filas y columnas para la hoja "QR", solo si existen columnas "QR"
                if columnas_qr:
                    df_qr = df_total[df_total[columnas_qr].sum(axis=1) != 0][
                        ["Fecha", "Liquidacion"] + columnas_qr
                    ]
                else:
                    df_qr = (
                        None  # Si no hay columnas "QR", no creamos el DataFrame "QR"
                    )

                # Identificar las columnas que contienen "AJUSTE"
                columnas_ajuste = [col for col in df_total.columns if "AJUSTE" in col]

                # Filtrar filas y columnas para la hoja "AJUSTE", solo si existen columnas "AJUSTE"
                if columnas_ajuste:
                    df_ajuste = df_total[df_total[columnas_ajuste].sum(axis=1) != 0][
                        ["Fecha", "Liquidacion"] + columnas_ajuste
                    ]
                else:
                    df_ajuste = None  # Si no hay columnas "AJUSTE", no creamos el DataFrame "AJUSTE"

                # Filtrar el DataFrame "Movimientos" eliminando las columnas "QR" y "AJUSTE"
                df_movimientos = df_total.drop(columns=columnas_qr + columnas_ajuste)

                # Ordenar las columnas para que "IMPORTE NETO" esté antes de las columnas que comienzan con "VENTAS"
                columnas_importe_neto = [
                    col for col in df_movimientos.columns if "IMPORTE NETO" in col
                ]
                columnas_ventas = [
                    col for col in df_movimientos.columns if col.startswith("VENTAS")
                ]
                columnas_restantes = [
                    col
                    for col in df_movimientos.columns
                    if col
                    not in columnas_importe_neto
                    + columnas_ventas
                    + ["Fecha", "Liquidacion"]
                ]

                nuevo_orden_columnas = (
                    ["Fecha", "Liquidacion"]
                    + columnas_restantes
                    + columnas_importe_neto
                    + columnas_ventas
                )

                df_movimientos = df_movimientos[nuevo_orden_columnas]

                # Obtener el directorio actual
                directorio_actual = os.getcwd()

                # Ruta del archivo Excel
                archivo_salida = os.path.join(directorio_actual, "Liquidaciones.xlsx")

                # Guardar en un archivo Excel con múltiples hojas
                with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
                    # Guardar la hoja "Movimientos"
                    df_movimientos.to_excel(
                        writer, sheet_name="Movimientos", index=False
                    )

                    # Guardar la hoja "QR" solo si tiene datos, o agregar un texto si no hay
                    if df_qr is not None:
                        df_qr.to_excel(writer, sheet_name="QR", index=False)
                    else:
                        wb = writer.book
                        ws_qr = wb.create_sheet("QR")
                        ws_qr["A1"] = "No hubo liquidaciones con QR"

                    # Guardar la hoja "AJUSTE" solo si tiene datos, o agregar un texto si no hay
                    if df_ajuste is not None:
                        df_ajuste.to_excel(writer, sheet_name="AJUSTE", index=False)
                    else:
                        wb = writer.book
                        ws_ajuste = wb.create_sheet("AJUSTE")
                        ws_ajuste["A1"] = "No hubo liquidaciones con Ajuste"

                # Abrir el archivo Excel con openpyxl para aplicar formato
                wb = openpyxl.load_workbook(archivo_salida)

                # Crear un objeto de estilo de borde fino
                border = Border(
                    left=Side(border_style="thin"),
                    right=Side(border_style="thin"),
                    top=Side(border_style="thin"),
                    bottom=Side(border_style="thin"),
                )

                def formatear_hoja(ws, df, columnas_ignorar):
                    for row in ws.iter_rows():
                        for cell in row:
                            if (
                                cell.column_letter not in columnas_ignorar
                            ):  # No aplicar formato en las columnas ignoradas
                                if isinstance(cell.value, (int, float)):
                                    cell.number_format = (
                                        "0.00"  # Formato de 2 decimales
                                    )
                            cell.border = border

                    # Agregar la fórmula de suma en la fila de totales si hay datos
                    total_row = len(df) + 2
                    for col_idx, column in enumerate(df.columns[2:], start=3):
                        ws.cell(
                            row=total_row,
                            column=col_idx,
                            value=f"=SUM({openpyxl.utils.get_column_letter(col_idx)}2:{openpyxl.utils.get_column_letter(col_idx)}{total_row-1})",
                        )

                # Formatear las hojas "Movimientos", "QR" y "AJUSTE"
                formatear_hoja(wb["Movimientos"], df_movimientos, ["A", "B"])
                if df_qr is not None:
                    formatear_hoja(wb["QR"], df_qr, ["A", "B"])
                if df_ajuste is not None:
                    formatear_hoja(wb["AJUSTE"], df_ajuste, ["A", "B"])

                # Guardar el archivo con el formato y totales aplicados
                wb.save(archivo_salida)

                return send_from_directory(
                    os.getcwd(), "Liquidaciones.xlsx", as_attachment=True
                )

            except Exception as e:
                return f"Hubo un error al procesar el archivo: {e}"

    return render_template("index.html")


if __name__ == "__main__":
    if not os.path.exists(app.config["UPLOAD_FOLDER"]):
        os.makedirs(app.config["UPLOAD_FOLDER"])
    app.run(debug=True, host="0.0.0.0", port=os.getenv("PORT", default=5000))
