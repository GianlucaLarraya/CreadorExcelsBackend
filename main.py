import os
from fastapi import FastAPI, File, UploadFile, Form, Request, Query, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import List
import cloudinary
import cloudinary.uploader
from dotenv import load_dotenv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request as GoogleRequest
import pickle
import base64
import requests


   # Solo si el archivo binario no existe ya
if os.path.exists("token.pickle.b64") and not os.path.exists("token.pickle"):
    with open("token.pickle.b64", "r") as f:
        b64data = f.read()
    with open("token.pickle", "wb") as f:
        f.write(base64.b64decode(b64data))

ENV = os.getenv("ENVIRONMENT", "development")

if ENV == "production":
    allowed_origins = ["https://creador-excels.vercel.app"]  # sin barra final
else:
    allowed_origins = ["*"]  

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Cargar variables de entorno
load_dotenv()

cloudinary.config(
    cloud_name=os.getenv('CLOUDINARY_CLOUD_NAME'),
    api_key=os.getenv('CLOUDINARY_API_KEY'),
    api_secret=os.getenv('CLOUDINARY_API_SECRET')
)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

API_TOKEN = os.getenv("API_TOKEN", "sorrento")

@app.middleware("http")
async def check_token(request: Request, call_next):
    # Permite el acceso a la documentación y a la raíz sin token
    if request.method == "OPTIONS":
        return await call_next(request)
    if request.url.path.startswith("/docs") or request.url.path.startswith("/openapi.json") or request.url.path == "/":
        return await call_next(request)
    # Verifica el header personalizado
    token = request.headers.get("x-api-token")
    if token != API_TOKEN:
        raise HTTPException(status_code=401, detail="Unauthorized")
    return await call_next(request)



# Estructura temporal para guardar prendas (en memoria)
pedido_prendas = []

@app.post("/agregar_prenda/")
async def agregar_prenda(
    foto: UploadFile = File(...),
    descripcion: str = Form(""),
    cantidades: str = Form(...),
    talles: str = Form(...),
):
    # Subir imagen a Cloudinary
    result = cloudinary.uploader.upload(foto.file, folder="pedidos")
    url_imagen = result["secure_url"]
    # Guardar prenda en la lista temporal
    prenda = {
        "url_imagen": url_imagen,
        "descripcion": descripcion,
        "cantidades": [int(x) if x else 0 for x in cantidades.split(",")],
        "talles": [t.strip() for t in talles.split(",")],
    }
    pedido_prendas.append(prenda)
    return JSONResponse({"ok": True, "url_imagen": url_imagen})

@app.get("/listar_prendas/")
def listar_prendas():
    return pedido_prendas

@app.post("/generar_google_sheet/")
async def generar_google_sheet(request: Request):
    try:
        data = await request.json()
        prendas = data.get('prendas', [])
        spreadsheet_id = data.get('spreadsheetId')  # <-- Nuevo parámetro opcional
        if not prendas:
            return JSONResponse({"ok": False, "msg": "No hay prendas cargadas"}, status_code=400)

        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(GoogleRequest())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials_oauth.json', SCOPES)
                creds = flow.run_local_server(port=8080)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)

        # Encabezados y valores
        talles = prendas[0]['talles']
        values = []
        for prenda in prendas:
            # Usar fórmula =IMAGE() para imágenes en columna separada
            fila = [f'=IMAGE("{prenda["url_imagen"]}")'] + prenda['cantidades']
            values.append(fila)

        if spreadsheet_id:
            # AGREGAR FILAS a hoja existente
            # Busca el nombre de la primera hoja
            sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet = sheet_metadata['sheets'][0]
            sheet_name = sheet['properties']['title']
            sheet_id = sheet['properties']['sheetId']
            # Encuentra la última fila con datos
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:A"
            ).execute()
            existing_rows = len(result.get('values', []))
            last_row = existing_rows + 1
            
            # Calcular totales de las nuevas prendas
            nuevas_prendas = sum(sum(int(cant) if cant else 0 for cant in prenda['cantidades']) for prenda in prendas)
            
            # Obtener total existente de la hoja
            total_existente = 0
            tiene_fila_totales = False
            if existing_rows > 0:
                # Buscar la fila de totales (última fila)
                total_result = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"{sheet_name}!A{existing_rows}:B{existing_rows}"
                ).execute()
                total_values = total_result.get('values', [])
                if total_values and total_values[0] and total_values[0][0] == "Total":
                    try:
                        total_existente = int(total_values[0][1]) if len(total_values[0]) > 1 else 0
                        tiene_fila_totales = True
                    except (ValueError, TypeError):
                        total_existente = 0
            
            # Calcular total acumulado
            total_prendas = total_existente + nuevas_prendas
            
            # Si la hoja está vacía, agregar encabezados primero
            headers = ["Imagen"] + [f"Talle ({t})" for t in talles]
            values_to_insert = []
            if existing_rows == 0:
                values_to_insert.append(headers)
            values_to_insert.extend(values)
            
            # Si ya había una fila de totales, la eliminamos antes de agregar las nuevas filas
            if tiene_fila_totales:
                # Eliminar la última fila (fila de totales)
                delete_request = {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": existing_rows - 1,
                            "endIndex": existing_rows
                        }
                    }
                }
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"requests": [delete_request]}
                ).execute()
                # Ajustar el número de filas existentes
                existing_rows -= 1
            
            # Agregar UNA SOLA fila de totales al final
            total_row = ["Total", total_prendas] + [""] * (len(talles) - 1)
            values_to_insert.append(total_row)
            
            # Agrega las filas nuevas (con o sin encabezados)
            service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A1",
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body={'values': values_to_insert}
            ).execute()
            # Ajustar tamaño de columna A (imágenes) y filas de imágenes (180px)
            requests = [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 1
                        },
                        "properties": {"pixelSize": 180},
                        "fields": "pixelSize"
                    }
                }
            ]
            # Ajustar alto de filas nuevas (todas menos encabezado si ya existía)
            start_row = existing_rows if existing_rows > 0 else 1  # Si ya había encabezado, solo las nuevas
            # Ajustar filas de datos (excluyendo la fila de totales)
            for i in range(start_row, start_row + len(values)):
                requests.append({
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": i,
                            "endIndex": i+1
                        },
                        "properties": {"pixelSize": 180},
                        "fields": "pixelSize"
                    }
                })
            # Centrar encabezados y cantidades por talle
            num_talles = len(talles)
            # Centrar encabezados de talles
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 1,
                        "endColumnIndex": 1 + num_talles
                    },
                    "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                    "fields": "userEnteredFormat.horizontalAlignment"
                }
            })
            # Centrar cantidades por talle (todas las filas de datos, excluyendo totales)
            if existing_rows == 0:
                data_start = 1
            else:
                data_start = existing_rows
            data_end = data_start + len(values)
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": data_start,
                        "endRowIndex": data_end,
                        "startColumnIndex": 1,
                        "endColumnIndex": 1 + num_talles
                    },
                    "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                    "fields": "userEnteredFormat.horizontalAlignment"
                }
            })
            
            # Formatear fila de totales (negrita y fondo gris)
            total_row_index = data_end  # La fila de totales va después de los datos
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": total_row_index,
                        "endRowIndex": total_row_index + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1 + num_talles
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                            "textFormat": {"bold": True},
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment"
                }
            })
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()
            
            url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'
        else:
            # CREAR NUEVA HOJA (como antes)
            headers = ["Imagen"] + [f"Talle ({t})" for t in talles]
            
            # Calcular totales
            total_prendas = sum(sum(int(cant) if cant else 0 for cant in prenda['cantidades']) for prenda in prendas)
            
            # Agregar UNA SOLA fila de totales al final
            total_row = ["Total", total_prendas] + [""] * (len(talles) - 1)
            all_values = [headers] + values + [total_row]
            sheet_title = data.get('sheetTitle', 'Pedido generado por API')
            spreadsheet = {
                'properties': {
                    'title': sheet_title
                }
            }
            spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
            spreadsheet_id = spreadsheet.get('spreadsheetId')
            # Escribir los datos
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='A1',
                valueInputOption='USER_ENTERED',
                body={'values': all_values}
            ).execute()
            # Ajustar tamaño de columna A (imágenes) y filas de imágenes (180px)
            requests = [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": 0,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 1
                        },
                        "properties": {"pixelSize": 180},
                        "fields": "pixelSize"
                    }
                }
            ]
            # Ajustar alto de filas con imágenes (todas menos encabezado)
            for i in range(1, len(all_values)):
                requests.append({
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": 0,
                            "dimension": "ROWS",
                            "startIndex": i,
                            "endIndex": i+1
                        },
                        "properties": {"pixelSize": 180},
                        "fields": "pixelSize"
                    }
                })
            # Centrar encabezados de talles
            num_talles = len(talles)
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 1,
                        "endColumnIndex": 1 + num_talles
                    },
                    "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                    "fields": "userEnteredFormat.horizontalAlignment"
                }
            })
            # Centrar cantidades por talle (todas las filas de datos)
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 1,
                        "endRowIndex": len(all_values) - 1,  # Excluir la fila de totales
                        "startColumnIndex": 1,
                        "endColumnIndex": 1 + num_talles
                    },
                    "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                    "fields": "userEnteredFormat.horizontalAlignment"
                }
            })
            
            # Formatear fila de totales (negrita y fondo gris)
            total_row_index = len(all_values) - 1
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": total_row_index,
                        "endRowIndex": total_row_index + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1 + num_talles
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                            "textFormat": {"bold": True},
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment"
                }
            })
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()
            
            url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'

        return JSONResponse({"ok": True, "url": url})
    except Exception as e:
        print(f"Error generando Google Sheet: {e}")
        return JSONResponse({"ok": False, "msg": f"Error generando la hoja: {str(e)}"}, status_code=500)

@app.get("/listar_sheets/")
def listar_sheets(force_refresh: bool = False):
    try:
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(GoogleRequest())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials_oauth.json', SCOPES)
                creds = flow.run_local_server(port=8080)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        
        # Parámetros más agresivos para evitar cache
        # Solo buscar en el drive principal del usuario, no en drives compartidos
        params = {
            "q": "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false and 'me' in owners",
            "pageSize": 100,  # Aumentar para obtener más resultados
            "fields": "files(id, name, webViewLink, modifiedTime, trashed, owners, permissions)",
            "orderBy": "modifiedTime desc",  # Siempre ordenar por fecha de modificación
            "corpora": "user",  # Solo archivos del usuario actual
            "includeItemsFromAllDrives": False  # No incluir drives compartidos
        }
        
        # Si se fuerza refresh, agregar parámetros adicionales
        if force_refresh:
            params["orderBy"] = "modifiedTime desc"
        
        results = drive_service.files().list(**params).execute()
        files = results.get('files', [])
        
        # Filtrar por nombre que comience exactamente con el string y NO esté en papelera
        filtered = [f for f in files if f['name'].startswith('Pedido') and not f.get('trashed', False)]
        
        print(f"Archivos encontrados inicialmente: {len(files)}")
        print(f"Después de filtrar por 'Pedido': {len(filtered)}")
        
        # Verificar que las hojas realmente existen y son accesibles
        accessible_files = []
        for f in filtered:
            try:
                # Verificar que el usuario actual es el propietario
                owners = f.get('owners', [])
                is_owner = any(owner.get('emailAddress') == creds.service_account_email if hasattr(creds, 'service_account_email') else True for owner in owners)
                
                if not is_owner:
                    print(f"  ⚠ {f['name']} (ID: {f['id']}) - NO ES PROPIETARIO")
                    continue
                
                # Intentar acceder a la hoja para verificar que existe
                sheets_service.spreadsheets().get(spreadsheetId=f['id']).execute()
                accessible_files.append(f)
                print(f"  ✓ {f['name']} (ID: {f['id']}) - ACCESIBLE Y PROPIETARIO")
            except Exception as e:
                print(f"  ✗ {f['name']} (ID: {f['id']}) - NO ACCESIBLE: {str(e)}")
        
        # Ordenar por fecha de modificación (más reciente primero)
        accessible_files.sort(key=lambda x: x.get('modifiedTime', ''), reverse=True)
        
        print(f"Encontradas {len(accessible_files)} hojas de pedidos accesibles y propias (de {len(filtered)} total)")
        
        return accessible_files
    except Exception as e:
        print(f"Error en listar_sheets: {e}")
        return [] 

@app.get("/leer_encabezados_sheet/")
def leer_encabezados_sheet(spreadsheet_id: str = Query(...)):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(GoogleRequest())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials_oauth.json', SCOPES)
            creds = flow.run_local_server(port=8080)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_name = sheet_metadata['sheets'][0]['properties']['title']
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1:Z1"
    ).execute()
    values = result.get('values', [])
    talles = values[0][1:] if values and len(values[0]) > 1 else []
    return {"talles": talles}

@app.post("/limpiar_cache/")
def limpiar_cache():
    """Elimina el token y fuerza una nueva autenticación"""
    try:
        if os.path.exists('token.pickle'):
            os.remove('token.pickle')
            print("Token eliminado. Se requerirá nueva autenticación.")
        if os.path.exists('token.pickle.b64'):
            os.remove('token.pickle.b64')
            print("Token base64 eliminado.")
        return JSONResponse({"ok": True, "msg": "Cache limpiado. Se requerirá nueva autenticación."})
    except Exception as e:
        return JSONResponse({"ok": False, "msg": f"Error limpiando cache: {str(e)}"}, status_code=500)

    