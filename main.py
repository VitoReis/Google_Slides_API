from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SLIDESSCOPES = ['https://www.googleapis.com/auth/presentations']
DRIVESCOPES = ['https://www.googleapis.com/auth/drive']


PRESENTATION_ID = 'SLIDEID'


def main():
    slidesCreds = None
    driveCreds = None

    # Check google slides credentials
    if os.path.exists('slidestoken.json'):
        slidesCreds = Credentials.from_authorized_user_file('slidestoken.json', SLIDESSCOPES)
    if not slidesCreds or not slidesCreds.valid:
        if slidesCreds and slidesCreds.expired and slidesCreds.refresh_token:
            slidesCreds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SLIDESSCOPES)
            slidesCreds = flow.run_local_server(port=0)
        with open('slidestoken.json', 'w') as token:
            token.write(slidesCreds.to_json())

    # Check google drive credentials
    if os.path.exists('drivetoken.json'):
        driveCreds = Credentials.from_authorized_user_file('drivetoken.json', DRIVESCOPES)
    if not driveCreds or not driveCreds.valid:
        if driveCreds and driveCreds.expired and driveCreds.refresh_token:
            driveCreds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', DRIVESCOPES)
            driveCreds = flow.run_local_server(port=0)
        with open('drivetoken.json', 'w') as token:
            token.write(driveCreds.to_json())

    try:
        slides_service = build('slides', 'v1', credentials=slidesCreds)
        drive_service = build('drive', 'v3', credentials=driveCreds)
        # Create a new slide title
        copy_title = 'Proposta comercial - Café Vertentes'
        body = {
            'name': copy_title
        }
        # Create a copy of the original slide on your drive
        drive_response = drive_service.files().copy(fileId=PRESENTATION_ID, body=body).execute()
        # Get the new slide ID
        presentation_copy_id = drive_response.get('id')
        # Set the requests
        requests = [
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{projeto}}',
                        'matchCase': True
                    },
                    'replaceText': 'Café Vertentes'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{descricao}}',
                        'matchCase': True
                    },
                    'replaceText': 'Levar grãos de alta qualidade para todos aqueles amantes de café que queiram desfrutar de uma experiência nova e com um café saboroso e com um preço justo.'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{problema}}',
                        'matchCase': True
                    },
                    'replaceText': 'Mercado ainda não consolidado na região.'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{solucao}}',
                        'matchCase': True
                    },
                    'replaceText': 'Entrega de produtos de alta qualidade a um preço acessível aos clientes.'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{Beneficio1}}',
                        'matchCase': True
                    },
                    'replaceText': 'Produtos de alta qualidade;'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{Beneficio2}}',
                        'matchCase': True
                    },
                    'replaceText': 'Preço acessível;'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{Beneficio3}}',
                        'matchCase': True
                    },
                    'replaceText': 'Foco na satisfação do cliente;'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{Beneficio4}}',
                        'matchCase': True
                    },
                    'replaceText': 'Alta produção.'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{custoDireto}}',
                        'matchCase': True
                    },
                    'replaceText': '1750,75'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{custoIndireto}}',
                        'matchCase': True
                    },
                    'replaceText': '900,00'
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{{plano}}',
                        'matchCase': True
                    },
                    'replaceText': 'Explorar um mercado ainda não consolidado na região e tornar-se referência.'
                }
            }
        ]

        body = {
            'requests': requests
        }
        # Execute the requests
        slides_service.presentations().batchUpdate(presentationId=presentation_copy_id, body=body).execute()

    except HttpError as error:
        print(f"An error occurred: {error}")
        print(error)


if __name__ == '__main__':
    main()

# ARRUMAR CREDENCIAL DO DRIVE