import os
from fastapi import FastAPI, File, UploadFile, Form, Request, Query
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

ENV = os.getenv("ENVIRONMENT", "development")

if ENV == "production":
    allowed_origins = ["https://creador-excels.vercel.app/"]  
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
        # Si la hoja está vacía, agregar encabezados primero
        headers = [f"Talle ({t})" for t in talles]
        values_to_insert = []
        if existing_rows == 0:
            values_to_insert.append(headers)
        values_to_insert.extend(values)
        # Agrega las filas nuevas (con o sin encabezados)
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': values_to_insert}
        ).execute()
        # Ajustar tamaño de columna A y filas de imágenes (180px)
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
        # Centrar cantidades por talle (todas las filas de datos)
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
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()
        url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'
    else:
        # CREAR NUEVA HOJA (como antes)
        headers = [f"Talle ({t})" for t in talles]
        all_values = [headers] + values
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
        # Ajustar tamaño de columna A y filas de imágenes (180px)
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
                    "endRowIndex": len(all_values),
                    "startColumnIndex": 1,
                    "endColumnIndex": 1 + num_talles
                },
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()
        url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'

    return JSONResponse({"ok": True, "url": url})

@app.get("/listar_sheets/")
def listar_sheets():
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
    results = drive_service.files().list(
        q="mimeType='application/vnd.google-apps.spreadsheet'",
        pageSize=50,
        fields="files(id, name, webViewLink)").execute()
    files = results.get('files', [])
    # Filtrar por nombre que comience exactamente con el string
    filtered = [f for f in files if f['name'].startswith('Pedido')]
    return filtered 

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

    