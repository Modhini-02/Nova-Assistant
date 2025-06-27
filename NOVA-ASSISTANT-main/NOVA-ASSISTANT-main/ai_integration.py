# ai_integration.py
import requests

GEMINI_API_KEY = "AIzaSyCPjwvy4Ib8zHmtD0JbMmuhc2mPSyO-cFs"
# Corrected Gemini endpoint
GEMINI_ENDPOINT = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key={GEMINI_API_KEY}"

def ai_response(prompt):
    try:
        headers = {"Content-Type": "application/json"}
        data = {
            "contents": [{
                "parts": [{"text": prompt}]
            }],
            "generationConfig": {
                "temperature": 0.9,
                "topK": 1,
                "topP": 1,
                "maxOutputTokens": 2048,
                "stopSequences": []
            },
            "safetySettings": [
                {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
            ]
        }

        response = requests.post(GEMINI_ENDPOINT, json=data, headers=headers)
        
        if response.status_code != 200:
            return f"API Error: {response.status_code} - {response.text}"

        response_json = response.json()
        
        # Extract response text
        candidates = response_json.get("candidates", [])
        if not candidates:
            return "No response generated."

        full_response = candidates[0].get("content", {}).get("parts", [{}])[0].get("text", "")
        return full_response.strip()[:2000]  # Increased character limit

    except Exception as e:
        return f"Error: {str(e)}"