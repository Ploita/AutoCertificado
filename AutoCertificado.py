"""
Código para automatizar a geração de certificados dado um modelo em Google Slides e uma planilha Google Sheets

Se aparecer a mensagem "Tudo Certo", significa que deu tudo certo
Se não apareceu a mensagem "Tudo certo", significa que deu errado

O código funciona utilizando API's do Google, para funcionar, precisa baixar a chave

Caso você não saiba fazer isso, vai ter o link no read.me
O vídeo ensina todo o passo-a-passo de como criar um projeto no Google Cloud também
"""
import os
import datetime
import time
import io
import pickle

from googleapiclient.http import MediaIoBaseUpload, MediaFileUpload, MediaIoBaseDownload
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from google.auth.transport.requests import Request

#A localização da chave
CLIENT_SECRET_FILE = r"path"
#ID do certificado
template_document_id = 'ID'
#ID da planilha com os dados
google_sheets_id = 'second ID'
#ID da pasta do evento (Vão ser criadas sub-pastas para organizar os certificados)
folder_id = 'folder ID' 

def Create_Service(client_secret_file, api_name, api_version, *scopes, prefix=''):
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    SCOPES = [scope for scope in scopes[0]]
    
    cred = None
    working_dir = os.getcwd()
    token_dir = 'token files'
    pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}{prefix}.pickle'

    ### Check if token dir exists first, if not, create the folder
    if not os.path.exists(os.path.join(working_dir, token_dir)):
        os.mkdir(os.path.join(working_dir, token_dir))

    if os.path.exists(os.path.join(working_dir, token_dir, pickle_file)):
        with open(os.path.join(working_dir, token_dir, pickle_file), 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            cred = flow.run_local_server()

        with open(os.path.join(working_dir, token_dir, pickle_file), 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, API_VERSION, 'service created successfully')
        return service
    except Exception as e:
        print(e)
        print(f'Failed to create service instance for {API_SERVICE_NAME}')
        os.remove(os.path.join(working_dir, token_dir, pickle_file))
        return None

def convert_to_RFC_datetime(year=1900, month=1, day=1, hour=0, minute=0):
    dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
    return dt

if __name__ == '__main__':
    API_NAME = 'calendar'
    API_VERSION = 'v3'
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    CLIENT_FILE = CLIENT_SECRET_FILE
    service = Create_Service(CLIENT_FILE, API_NAME, API_VERSION, SCOPES, 'x')

"""
Step 1. Create Google API Service Instances
"""
# Google Slides instance 
service_docs = Create_Service(
    CLIENT_SECRET_FILE, 
    'slides', 'v1', 
    ['https://www.googleapis.com/auth/presentations']
    
)
time.sleep(2)

# Google Drive instance
service_drive = Create_Service(
    CLIENT_SECRET_FILE,
    'drive',
    'v3',
    ['https://www.googleapis.com/auth/drive']
)
time.sleep(2)

# Google Sheets instance
service_sheets = Create_Service(
    CLIENT_SECRET_FILE,
    'sheets',
    'v4',
    ['https://www.googleapis.com/auth/spreadsheets']
)
time.sleep(2)


responses = {}

"""
Step 2. Load Records from Google Sheets
"""

worksheet_name = 'Projects'
responses['sheets'] = service_sheets.spreadsheets().values().get(
    spreadsheetId=google_sheets_id,
    range=worksheet_name,
    majorDimension='ROWS',
).execute()

columns = responses['sheets']['values'][0]
records = responses['sheets']['values'][1:]

def create_folder_in_folder(folder_name,parent_folder_id):
    
    file_metadata = {
    'name' : folder_name,
    'parents' : [folder_id],
    'mimeType' : 'application/vnd.google-apps.folder'
    }

    file = service_drive.files().create(body=file_metadata,
                                    fields='id').execute()
    
    return file.get('id')


"""
Step 3. Iterate Each Record and Perform Mail Merge
"""
page = service_docs.presentations().pages().get
def mapping(merge_field, value=''):
    json_representation = {
        "replaceAllText": {
            "replaceText": value,
            "pageObjectIds": [],
            "containsText": {
                'text': '{{{{{0}}}}}'.format(merge_field),
                'matchCase': 'true'        
            },
        }
    }
    return json_representation

slide_folder = create_folder_in_folder('Certificados_Slide',folder_id)
pdf_folder = create_folder_in_folder('Certificados',folder_id)

for val,record in enumerate(records):
    print('Processing record {0}...'.format(val+1))
    # Copy template doc file as new doc file
    document_title = '{0}'.format(record[0])

    responses['docs'] = service_drive.files().copy(
        fileId=template_document_id,
        body={
            'parents': [slide_folder],
            'name': document_title
        }
        
    ).execute()
    
    document_id = responses['docs']['id']
    
    # Update Google Docs document (not template file)
    merge_fields_information = [mapping(columns[indx], value) for indx, value in enumerate(record)]

    service_docs.presentations().batchUpdate(
        presentationId = document_id,
        body={
            'requests': merge_fields_information
        }
    ).execute()
    
    """"
    Export Document as PDF
    """
    
    PDF_MIME_TYPE = 'application/pdf'
    byteString = service_drive.files().export(  
        fileId=document_id,
        mimeType=PDF_MIME_TYPE
    ).execute()

    media_object = MediaIoBaseUpload(io.BytesIO(byteString), mimetype=PDF_MIME_TYPE)
    service_drive.files().create(
        media_body=media_object,
        body={
            'parents': [pdf_folder],
            'name': '{0} (PDF).pdf'.format(document_title)
        }
    ).execute()
    
print('Tudo Certo')