from flask import Flask, render_template, request, send_from_directory
import os
import re
import pandas as pd

app = Flask(__name__)


# Ruta para la carga de archivos y procesamiento
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["txt_file"]
        if file:

            if not os.path.exists("uploads"):
                os.makedirs("uploads")

            file_path = os.path.join("uploads", file.filename)
            file.save(file_path)

            try:
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        lines = f.readlines()
                except UnicodeDecodeError:
                    with open(file_path, "r", encoding="latin-1") as f:
                        lines = f.readlines()

                # Limpiar caracteres de control y eliminar bloques entre "----" y "--"
                cleaned_lines = []
                movements = []
                eliminar = False
                eliminar_desde_totales = False
                temp_movement = {}

                for i, line in enumerate(lines[9:], start=2):
                    # Si la línea contiene "TOTALES POR TASA", eliminar todo lo posterior
                    if "TOTALES POR TASA" in line:
                        eliminar_desde_totales = True
                        continue  # Saltar esta línea (no la procesamos, solo marcamos la eliminación)

                    if eliminar_desde_totales:
                        break  # Detener el procesamiento de líneas después de "TOTALES POR TASA"

                    # Si la línea comienza con "----", marca el inicio del bloque a eliminar
                    if line.startswith("----"):
                        eliminar = True
                        continue  # Saltar esta línea (ya no se procesa)

                    # Si la línea comienza con "--", marca el final del bloque a eliminar
                    if line.startswith("--"):
                        eliminar = False
                        continue  # Saltar esta línea (ya no se procesa)

                    # Si no estamos dentro de un bloque a eliminar, limpiamos la línea
                    if not eliminar:
                        cleaned_line = re.sub(
                            r"\x1b[^m]*m", "", line
                        )  # Elimina secuencias ANSI
                        cleaned_line = re.sub(
                            r"[\x00-\x1F\x7F]", "", cleaned_line
                        )  # Eliminación de caracteres de control ASCII
                        cleaned_lines.append(
                            cleaned_line
                        )  # Guardar la línea limpia sin espacios extra

                doble_cleaned_lines = []

                for i, line in enumerate(cleaned_lines, start=-1):
                    if "PPag." in line:
                        break
                    doble_cleaned_lines.append(line)

                for index, cleaned_line in enumerate(doble_cleaned_lines):
                    if (
                        "numero" in temp_movement
                        and temp_movement["numero"] == cleaned_line[12:20]
                    ):
                        partes = re.split(r"\s{3,}", cleaned_line[70:])
                        if len(partes) < 2:
                            pass
                        else:
                            if partes[0] == "Tasa 21%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                            elif partes[0] == "T.10.5%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                            elif partes[0] == "Tasa 27%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                            elif partes[0] == "C.F.21%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                            elif partes[0] == "C.F.10.5%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                            elif partes[0] == "Tasa 2.5%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                            elif partes[0] == "T.IMP 21%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                            elif partes[0] == "T.IMP 10%":
                                temp_movement[partes[0] + " Neto"] = partes[1]
                                temp_movement[partes[0] + " IVA"] = partes[2]
                    else:
                        if cleaned_line[0:2] == "  ":
                            partes = re.split(r"\s{3,}", cleaned_line[70:])
                            if len(partes) < 2:
                                pass
                            else:
                                if partes[0] == "Tasa 21%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "T.10.5%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "Tasa 27%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "C.F.21%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "C.F.10.5%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "Tasa 2.5%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "T.IMP 21%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "T.IMP 10%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                else:
                                    temp_movement[partes[0]] = partes[1]
                            if (
                                index == len(doble_cleaned_lines) - 1
                            ):  # Verificar si es el último elemento
                                movements.append(temp_movement)
                        else:
                            partes = re.split(r"\s{3,}", cleaned_line[70:])
                            movement = temp_movement.copy()
                            movements.append(movement)
                            temp_movement.clear()
                            if len(partes) < 2:
                                pass
                            else:
                                temp_movement = {
                                    "Fecha": cleaned_line[0:2],
                                    "Comprobante": cleaned_line[3:5],
                                    "PV": cleaned_line[6:11],
                                    "Nro": cleaned_line[12:20],
                                    "Letra": cleaned_line[20:21],
                                    "Razon Social": cleaned_line[22:44],
                                    "Condicion": cleaned_line[45:49],
                                    "CUIT": cleaned_line[50:63],
                                    "Concepto": cleaned_line[64:67],
                                    "Jurisdiccion": cleaned_line[68:69],
                                }
                                if partes[0] == "Tasa 21%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "T.10.5%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "Tasa 27%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "C.F.21%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "C.F.10.5%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "Tasa 2.5%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "T.IMP 21%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                elif partes[0] == "T.IMP 10%":
                                    temp_movement[partes[0] + " Neto"] = partes[1]
                                    temp_movement[partes[0] + " IVA"] = partes[2]
                                else:
                                    temp_movement[partes[0]] = partes[1]
                                if (
                                    index == len(doble_cleaned_lines) - 1
                                ):  # Verificar si es el último elemento
                                    movements.append(temp_movement)

                if not movements[0]:
                    movements.pop(0)

                # Crear el DataFrame
                df = pd.DataFrame(movements)
                # Reemplazar los valores NaN por 0
                df = df.fillna(0)

                # Reemplazar comas por puntos en las columnas numéricas (desde el índice 11 hasta el final)
                df.iloc[:, 10:] = df.iloc[:, 10:].replace(",", ".", regex=True)

                # Convertir las columnas desde el índice 11 hasta el final a tipo numérico
                df.iloc[:, 10:] = (
                    df.iloc[:, 10:].apply(pd.to_numeric, errors="coerce").fillna(0)
                )

                # Lista de columnas que deseas volver negativas
                columnas_a_convertir = df.columns[10:]

                # Aplicar la conversión a negativos si el tipo de comprobante es "NC"
                df.loc[df["Comprobante"] == "NC", columnas_a_convertir] *= -1

                df["PV"] = pd.to_numeric(df["PV"])
                df["Nro"] = pd.to_numeric(df["Nro"])
                df["Concepto"] = pd.to_numeric(df["Concepto"])

                # Crear una lista para almacenar las filas combinadas
                resultado = []

                # Variable para almacenar la fila combinada actual
                fila_actual = df.iloc[0].copy()

                for i in range(1, len(df)):
                    fila_siguiente = df.iloc[i]

                    # Si la clave principal se repite, combinar valores
                    if (
                        fila_actual["Nro"] == fila_siguiente["Nro"]
                        and fila_actual["PV"] == fila_siguiente["PV"]
                        and fila_actual["Razon Social"]
                        == fila_siguiente["Razon Social"]
                    ):
                        for col in df.columns[10:]:  # Sumar solo las columnas numéricas
                            fila_actual[col] += fila_siguiente[col]
                    else:
                        resultado.append(fila_actual)
                        fila_actual = fila_siguiente.copy()

                # Agregar la última fila combinada
                resultado.append(fila_actual)

                # Convertir lista a DataFrame
                df_final = pd.DataFrame(resultado)

                df_final["Total"] = df_final.iloc[:, 10:].sum(axis=1)

                # Crear una fila con la etiqueta "TOTAL"
                fila_total = pd.DataFrame(df_final.iloc[:, 10:].sum()).T

                # Agregar un identificador en la primera columna para que se distinga la fila de totalización
                fila_total.insert(0, "Nro", "TOTAL")
                fila_total.insert(1, "Razon Social", "")

                # Concatenar la fila total con el DataFrame
                df_final = pd.concat([df_final, fila_total], ignore_index=True)

                # Guardar el DataFrame en un archivo Excel
                excel_filename = "Movimientos_MENDEZ.xlsx"
                df_final.to_excel(excel_filename, index=False)

                # Enviar el archivo Excel generado
                return send_from_directory(
                    os.getcwd(), excel_filename, as_attachment=True
                )

            except Exception as e:
                return f"Hubo un error al procesar el archivo: {e}"

    return render_template("index.html")


if __name__ == "__main__":
    # if not os.path.exists(app.config["UPLOAD_FOLDER"]):
    #     os.makedirs(app.config["UPLOAD_FOLDER"])
    app.run(debug=True, host="0.0.0.0", port=os.getenv("PORT", default=5000))
