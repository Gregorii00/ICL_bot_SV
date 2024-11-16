import csv
import io
import urllib.request

id_docs = '1Sft6GGeylsaeeD4QOZR2PJj0m4BhIcvOUuiRmuuV8sc'
url = f'https://docs.google.com/document/d/{id_docs}/export?format=csv'

response = urllib.request.urlopen(url)

with io.TextIOWrapper(response, encoding='utf-8') as f:
    reader = csv.reader(f)

    for row in reader:
        print(row)
