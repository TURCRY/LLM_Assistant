import requests
import json

url = "http://192.168.0.155:5050/annoter"
headers = {
    "x-api-key": "hy^wQ#4d3HpnEl4x1Mg&"
}

payload = {
    "prompt": "Qui était Victor Hugo ?",
    "temperature": 0.5,
    "top_p": 0.8,
    "top_k": 20,
    "repeat_penalty": 1.1,
    "max_tokens": 300,
    "user": "nicolas.turcry@ntu-consult.com"
}

try:
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    data = response.json()
    print("✅ Réponse du serveur :")
    print(json.dumps(data, indent=2, ensure_ascii=False))
except requests.exceptions.HTTPError as err:
    print(f"❌ Erreur HTTP : {err.response.status_code}")
    print(err.response.text)
except Exception as e:
    print(f"❌ Autre erreur : {e}")
