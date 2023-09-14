from datetime import datetime
import openpyxl
import pyodbc
import re

def main():
  # Obtenemos el archivo Excel
  excel_file = "LQR Database - Username MG.xlsx"

  # Obtenemos el nombre de la hoja de trabajo.
  sheet_name = "query"

  # Generamos la consulta de inserción.
  query = generate_insert_queries(excel_file, sheet_name)

  # Imprimimos la consulta.
  print(query)

def generate_insert_queries(excel_file, sheet_name):
  """
  Genera una consulta de inserción a partir de un archivo Excel.

  Args:
    excel_file: El archivo Excel.
    sheet_name: El nombre de la hoja de trabajo.

  Returns:
    La consultas de inserción.
  """

  # Abrimos el archivo Excel.
  workbook = openpyxl.load_workbook(excel_file)

  # Obtenemos la hoja de trabajo "Sheet1".
  worksheet = workbook[sheet_name]

  # Recorremos las filas del excel para generar cada uno de los diccionarios de valores
  rows_insert_values: list[dict] = []
  for indx, row in enumerate(worksheet.rows):
    if (
        # si es la cabecera...
        indx == 0 
        # si no hay usuario valido registrado...
        or re.match(r"^([A-Z\s]+)\s+\(([^\\]+)\\([A-Z0-9]+)\)$", str(row[3].value).strip()) == None 
        # si no es un usuario CI o LQR
        or ("CI" not in str(row[1].value) and "LQR" not in str(row[1].value))
      ):
      continue
    rows_insert_values.append(generate_insert_values(row))

  # Obtenemos los usuarios TLN
  tln_users = get_user_ids()
  # Obtenemos los ids de workshop
  get_workshop_ids([])

  # Recorremos la lista de valores obtenidos para generar el [workshop_id] y [user_id]
  for indx, insert_values in enumerate(rows_insert_values):
    # obtenemos el [user_id]
    user_id = tln_users.get(str(insert_values["user"]))
    if(user_id is None):
      rows_insert_values[indx]["is_valid_row"] = False
      continue
    rows_insert_values[indx]["is_valid_row"] = True
    rows_insert_values[indx]["user_id"] = user_id
    del rows_insert_values[indx]["user"]



def generate_insert_values(row):
  """
  Genera una consulta de inserción a partir de una fila de Excel.

  Args:
    row: Fila de la hoja de Excel.
  Returns:
    La consulta de inserción.
  """
  
  # Creamos una cadena de consulta vacía.
  insert_values = {
    "type": None,
    "workshop_legacy_plant_code": None,
    "user": None,
    "user_mail": None,
    "user_tel1": None,
    "user_tel2": None,
    "lqsm_score": None,
    "ph_score": None,
    "sr_score": None,
    "bo_score": None,
    "ab_score": None,
    "cp_score": None,
    "dl_score": None,
    "sca_score": None,
    "ex_score": None,
    "final_score": None,
    "approved_flag": None,
    "approved_by": None,
    "approved_date": None,
    "created_by": "SysAdmin",
    "created_date": datetime.utcnow(),
    "updated_by": None,
    "updated_date": None
  }

  # Agregamos los valores de la fila de datos.
  for cell in row:
    if cell.column_letter == "A":
      # El valor de la columna tiene la forma "316 - Erbil Pioneer Engineering"
      # Necesitamos el número del principio (code)
      insert_values["workshop_legacy_plant_code"] = int(str(cell.value).split("-")[0].strip()) 
    if cell.column_letter == "B":
      insert_values["type"] = "CI" if "CI" in str(cell.value) else "LQR"
    if cell.column_letter == "D":
      # El valor de la columna tiene la forma "KK SHANIL (SSN\SKK)"
      # Necesitamos el valor dentro de los parentesis (solo puede haber uno, asi que se obtiene el primero)
      insert_values["user"] = get_values_in_parenthesis(str(cell.value))[0]
    if cell.column_letter == "E":
      insert_values["user_mail"] = str(cell.value)
    if cell.column_letter == "F":
      insert_values["user_tel1"] = str(cell.value)
    if cell.column_letter == "G":
      insert_values["user_tel2"] = str(cell.value)
    if cell.column_letter == "H" and insert_values["type"] == "LQR":
      insert_values["lqsm_score"] = float(cell.value or 0)
    if cell.column_letter == "I" and insert_values["type"] == "LQR":
      insert_values["ph_score"] = float(cell.value or 0)
    if cell.column_letter == "J" and insert_values["type"] == "LQR":
      insert_values["sr_score"] = float(cell.value or 0)
    if cell.column_letter == "K" and insert_values["type"] == "LQR":
      insert_values["bo_score"] = float(cell.value or 0)
    if cell.column_letter == "L" and insert_values["type"] == "LQR":
      insert_values["ab_score"] = float(cell.value or 0)
    if cell.column_letter == "M" and insert_values["type"] == "LQR":
      insert_values["cp_score"] = float(cell.value or 0)
    if cell.column_letter == "N" and insert_values["type"] == "LQR":
      insert_values["dl_score"] = float(cell.value or 0)
    if cell.column_letter == "O" and insert_values["type"] == "LQR":
      insert_values["sca_score"] = float(cell.value or 0)
    if cell.column_letter == "P" and insert_values["type"] == "LQR":
      insert_values["ex_score"] = float(cell.value or 0)
    if cell.column_letter == "Q" and insert_values["type"] == "LQR":
      # Obtenemos cuantos examenes han registrados para este registro
      existing_scores: int = (1 if insert_values["lqsm_score"] else 0) + \
                             (1 if insert_values["ph_score"] else 0) + \
                             (1 if insert_values["sr_score"] else 0) + \
                             (1 if insert_values["bo_score"] else 0) + \
                             (1 if insert_values["ab_score"] else 0) + \
                             (1 if insert_values["cp_score"] else 0) + \
                             (1 if insert_values["dl_score"] else 0) + \
                             (1 if insert_values["sca_score"] else 0) + \
                             (1 if insert_values["ex_score"] else 0)
      # Obtenemos el score de los examenes que han registrados para este registro
      scores: float = (insert_values["lqsm_score"] if insert_values["lqsm_score"] else 0) + \
                      (insert_values["ph_score"] if insert_values["ph_score"] else 0) + \
                      (insert_values["sr_score"] if insert_values["sr_score"] else 0) + \
                      (insert_values["bo_score"] if insert_values["bo_score"] else 0) + \
                      (insert_values["ab_score"] if insert_values["ab_score"] else 0) + \
                      (insert_values["cp_score"] if insert_values["cp_score"] else 0) + \
                      (insert_values["dl_score"] if insert_values["dl_score"] else 0) + \
                      (insert_values["sca_score"] if insert_values["sca_score"] else 0) + \
                      (insert_values["ex_score"] if insert_values["ex_score"] else 0)
      # Hay un valor registrado para esta columna en la celda "D"
      # Sin embargo, para asegurar la calidad del dato, se calcula...
      final_score = float(scores/existing_scores) if existing_scores > 0 else 0
      approved_flag = True if final_score >= 70 else False
      insert_values["final_score"] = final_score
      insert_values["approved_flag"] = approved_flag
      if (approved_flag):
        insert_values["approved_by"] = "SysAdmin"
        insert_values["approved_date"] = datetime.utcnow()
    if cell.column_letter == "S" and insert_values["type"] == "CI":
      result = str(cell.value).strip().lower()

      approved_flag = None
      if (result == "pass"):
        approved_flag = True
      if (result == "fail"):
        approved_flag = False

      insert_values["approved_flag"] = approved_flag
      if (approved_flag):
        insert_values["approved_by"] = "SysAdmin"
        insert_values["approved_date"] = datetime.utcnow()
    if cell.column_letter == "T" and cell.value:
      insert_values["approved_date"] = cell.value
    if cell.column_letter == "AB":
      insert_values["updated_date"] = cell.value
    if cell.column_letter == "AC":
      insert_values["updated_by"] = str(cell.value)
  
  return insert_values

def get_values_in_parenthesis(text):
  """
  Obtiene los valores dentro de los paréntesis de una cadena de texto.

  Args:
    text: La cadena de texto.

  Returns:
    Una lista con los valores dentro de los paréntesis.
  """
  # Obtenemos los valores coincidentes.
  matches:list[str] = re.findall(r"\(.*?\)", text)

  for indx, match in enumerate(matches):
    matches[indx] = matches[indx].replace('(', '')
    matches[indx] = matches[indx].replace(')', '')

  return matches

def get_user_ids():
  """
  Obtiene el ID de usuario a partir de la tabla "Tln_users_PROD.xlsx".

  Args:
    tln_user: El nombre de usuario y dominio del usuario.

  Returns:
    El ID de usuario.
  """

  # Abrimos el archivo Excel.
  workbook = openpyxl.load_workbook("Tln_users_PROD.xlsx")

  # Obtenemos la hoja de trabajo "Sheet1".
  worksheet = workbook["Sheet1"]

  # Buscamos el registro que coincida con el nombre de usuario y dominio.
  output:dict = {}
  for indx, row in enumerate(worksheet.rows):
    if indx == 0:
      continue
    output[f"{row[2].value}\\{row[1].value}"] = row[0].value

  return output

def get_workshop_ids(workshop_legacy_plant_codes:list[int]):
  SERVER = 'TNSWDB39\STDQA'
  DATABASE = 'TLN_INC'
  # Crear la cadena de conexión# Creamos la cadena de conexión.
  connection_string = f"DRIVER=SQL Server;SERVER={SERVER};DATABASE={DATABASE};integrated security=SSPI;"

  # Conectarse a la base de datos
  conexion = pyodbc.connect(connection_string)

  # Crear un objeto cursor y ejecutar una consulta, si es necesario
  cursor = conexion.cursor()
  consulta = "select workshop_id, workshop_legacy_plant_code from tln_workshops"
  cursor.execute(consulta)
  resultados = cursor.fetchall()

  # Procesar los resultados, si es necesario
  for fila in resultados:
      columna1, columna2 = fila
      print(f"columna1: {columna1}, columna2: {columna2}")

  # Cerrar el cursor y la conexión, si es necesario
  cursor.close()
  conexion.close()

if __name__ == "__main__":
  main()
