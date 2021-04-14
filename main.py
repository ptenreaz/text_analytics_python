# Author: Jeffrey Prado

from azure.ai.textanalytics import TextAnalyticsClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd
import matplotlib.pyplot as plt

KEY = "<paste-your-text-analytics-key-here>"
ENDPOINT = "<paste-your-text-analytics-endpoint-here>"

def open_excel(archivo):
    """
        Lee el archivo con pandas

            Parámetros:
                archivo - Path del archivo a leer

            Retorna: Dataframe del archivo
    """
    return pd.read_excel(archivo)

def authenticate_client(key=KEY, endpoint=ENDPOINT):
    """
        Realiza la conexión con Azure Text Analytics

            Parámetros:
                key - API key del servicio en Azure
                endpoint - Endpoint del servicio creado

            Retorna: Cliente TextAnalyticsClient

            Excepciones:
                Exception: Excepción general con mensaje 
                incluido SI se presenta alguna en el request
    """
    try:
        ta_credential = AzureKeyCredential(key)
        text_analytics_client = TextAnalyticsClient(
                endpoint=endpoint, 
                credential=ta_credential)
        return text_analytics_client
    except Exception as err:
        print("Encountered exception. {}".format(err))

def analyze_comment(client, txt):
    """
        Analiza los sentimientos del texto txt

            Parametros:
                client: Cliente TextAnalyticsClient
                txt: Lista con texto a analizar

            Retorna: 
                result: Diccionario con valores por 
                cada sentimiento y el sentimiento total

            Excepciones:
                Exception: Excepción general con mensaje 
                incluido SI se presenta alguna en el request
    """
    try:
        result = client.analyze_sentiment([txt], show_opinion_mining=True)[0]
        result = {
            'neg': result.confidence_scores.negative,
            'neu': result.confidence_scores.neutral,
            'pos': result.confidence_scores.positive,
            'sent': result.sentences[0].sentiment
        }
        return result
    except Exception as err:
        print("Encountered exception. {}".format(err))

def detect_language(client, txt):
    """
        Detecta en idioma del texto txt

            Parametros:
                client: Cliente TextAnalyticsClient
                txt: Lista con texto para detectar idioma

            Retorna: 
                result: Diccionario con el idioma detectado

            Excepciones:
                Exception: Excepción general con mensaje 
                incluido SI se presenta alguna en el request
    """
    try:
        response = client.detect_language(documents=[txt])[0]
        return {'idio': response.primary_language.name}

    except Exception as err:
        print("Encountered exception. {}".format(err))

if __name__ == "__main__":
    df = open_excel(r"comentarios.xlsx")

    print('\nDataframe inicial')
    print(df)

    client = authenticate_client()

    print('\nAnalizando sentimientos...\n')
    df['Idioma']        = df.apply(lambda row: detect_language(client, row['Comentario']).get('idio', ''), axis=1)
    df['Negativo']      = df.apply(lambda row: analyze_comment(client, row['Comentario']).get('neg', ''), axis=1)
    df['Neutral']       = df.apply(lambda row: analyze_comment(client, row['Comentario']).get('neu', ''), axis=1)
    df['Positivo']      = df.apply(lambda row: analyze_comment(client, row['Comentario']).get('pos', ''), axis=1)
    df['Sentimiento']   = df.apply(lambda row: analyze_comment(client, row['Comentario']).get('sent', ''), axis=1)

    print(df)

    print('\nGenerando gráfico...\n')
    lista = list(df['Sentimiento'])
    values = [lista.count('neutral'), lista.count('negative'), lista.count('positive')]
    labels = ["Comentarios Neutrales", "Comentarios Negativos", "Comentarios Positivos"]
    plt.pie(values, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)
    plt.title("Distribución de Comentarios")
    plt.show()
    