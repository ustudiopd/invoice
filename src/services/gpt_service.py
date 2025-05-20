import requests


def ask_gpt_api(messages, api_key, model):
    """GPT API를 호출하여 응답을 받아옵니다."""
    if not api_key:
        return "[OpenAI API 키를 .env에 입력하세요]"
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "model": model,
        "messages": messages,
        "max_tokens": 2048,
        "temperature": 0.7
    }
    try:
        resp = requests.post(url, headers=headers, json=data, timeout=30)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"[GPT 호출 오류] {e}"