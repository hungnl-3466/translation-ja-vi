import requests

endpoint = 'https://doctranstesting.cognitiveservices.azure.com/'
key =  '81524c4f50fd44748bb9ceed9101859b'
path = 'translator/text/batch/v1.1/batches'
constructed_url = endpoint + path

sourceSASUrl = '{your-source-container-SAS-URL}'

targetSASUrl = '{your-target-container-SAS-URL}'

body= {
    "inputs": [
        {
            "source": {
                "sourceUrl": sourceSASUrl,
                "storageSource": "AzureBlob",
                "language": "en"
            },
            "targets": [
                {
                    "targetUrl": targetSASUrl,
                    "storageSource": "AzureBlob",
                    "category": "general",
                    "language": "es"
                }
            ]
        }
    ]
}
headers = {
  'Ocp-Apim-Subscription-Key': key,
  'Content-Type': 'application/json',
}

response = requests.post(constructed_url, headers=headers, json=body)
response_headers = response.headers

print(f'response status code: {response.status_code}\nresponse status: {response.reason}\n\nresponse headers:\n')

for key, value in response_headers.items():
    print(key, ":", value)