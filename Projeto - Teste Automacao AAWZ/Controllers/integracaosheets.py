import os.path
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1EnlMi7ZK3Q7qe4E5rbKD3CLogqdnOC33EssngAMQ7H0"
SAMPLE_RANGE_NAME = "Vendas!A1:J35"

def main():

  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  # If there are no (valid) credentials available, let the user log in.

  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)

    # Save the credentials for the next run

    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Ler informações da Planilha
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId="1EnlMi7ZK3Q7qe4E5rbKD3CLogqdnOC33EssngAMQ7H0", range="Vendas!A1:F35")
        .execute()
    )

    #Adicionando as informações
    df = pd.DataFrame(result[1:], columns=result[0])
    df['Valor da Venda'] = df['Valor da Venda'].replace({'R\\$ ': '', ',': '.'}, regex=True).astype(float)
    df['Custo da Venda'] = df['Custo da Venda'].replace({'R\\$ ': '', ',': '.'}, regex=True).astype(float)
    df['Comissão'] = df['Valor da Venda'] * 0.10
    df['Comissão Marketing'] = df.apply(lambda row: row['Comissão'] * 0.20 if row['Canal de Venda'] == 'Online' else 0, axis=1)
    df['Comissão Gerente'] = df.apply(lambda row: (row['Comissão'] - row['Comissão Marketing']) * 0.10 if (row['Comissão'] - row['Comissão Marketing']) >= 1500 else 0, axis=1)
    df['Comissão Líquida'] = df['Comissão'] - df['Comissão Marketing'] - df['Comissão Gerente']
    
    # Agrupar por vendedor
    comissoes = df.groupby('Nome do Vendedor').agg({
        'Comissão': 'sum',
        'Comissão Marketing': 'sum',
        'Comissão Gerente': 'sum',
        'Comissão Líquida': 'sum'
}).reset_index()

    # Renomear colunas para o resultado final
    comissoes.columns = ['Nome do Vendedor', 'Comissão Total', 'Para Marketing', 'Para Gerente', 'Valor Pago']

    # Escrever os resultados em uma nova aba
    result_sheet = sheet.add_worksheet(title="Comissões", rows="100", cols="20")
    result_sheet.update([comissoes.columns.values.tolist()] + comissoes.values.tolist())

    print("Cálculo de comissões concluído e resultados atualizados na aba 'Comissões'.")

  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()
  