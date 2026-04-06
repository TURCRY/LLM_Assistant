import requests
import json

url = "http://192.168.0.155:5050/annoter"
headers = {
    "x-api-key": "hy^wQ#4d3HpnEl4x1Mg&",
    "Content-Type": "application/json"
}

payload = {
    "prompt": "Quel est le rôle d’un État moderne ?",
    "temperature": 0.7,
    "top_p": 0.9,
    "top_k": 40,
    "repeat_penalty": 1.1,
    "max_tokens": 200,
    "system": "Vous êtes un assistant juridique."
}

print("📤 Envoi du prompt...")
response = requests.post(url, headers=headers, json=payload)

print("📥 Statut :", response.status_code)
print("📥 Réponse :")
print(response.text)
