import requests

openrouter_api_key = "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"
url = "https://openrouter.ai/api/v1/chat/completions"
headers = {
    "Authorization": f"Bearer {openrouter_api_key}",
    "Content-Type": "application/json"
}
data = {
    "model": "openai/gpt-4o",
    "messages": [
        {"role": "user", "content": "Bonjour ! Donne une phrase très courte de feedback sur une formation."}
    ]
}
r = requests.post(url, headers=headers, json=data)
print(r.status_code)
print(r.text)
