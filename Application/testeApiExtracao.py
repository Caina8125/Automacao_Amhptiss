import json
import time
import requests
import base64


def pdf_to_base64(file_path):
    with open(file_path, "rb") as pdf_file:
        encoded_pdf = base64.b64encode(pdf_file.read()).decode('utf-8')
    return encoded_pdf

def apiExtracao():

    file_path = r"C:\Users\lucas.timoteo\Desktop\56058403.pdf"
    base64_string = pdf_to_base64(file_path)


    proxies = {
        "http": "http://lucas.timoteo:81259018@10.0.0.230:3128/",
        "https": "http://lucas.timoteo:81259018@10.0.0.230:3128/"
    }

    url = 'http://54.87.118.248/process_document'

    headers = {
        'Content-Type': 'application/json'
    }

    data = {
        "Type": 1,
        "Carrier": 225,
        "Process": -864544,
        "pdf": base64_string
    }

    print(base64_string)

    post = requests.post(url, headers, data, proxies=proxies)
    print(post.status_code)

apiExtracao()