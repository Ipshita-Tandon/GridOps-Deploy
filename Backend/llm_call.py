import requests
import json
import os

api_key = os.getenv("DMG_API_KEY")


def make_llm_call(prompt):
    url = "https://dmg-stg.dcai.corp.adobe.com/chat/completions"
    headers = {
        "api-key": api_key,
        "Content-Type": "application/json"
    }
    data = {
        "model": "gpt-4o-mini",
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ],
        "max_tokens": 4000
    }
    
    response = requests.post(
        url,
        headers=headers,
        json=data
    )
    
    response_json = response.json()
    
    # Extract token usage information and write to JSON file
    if 'usage' in response_json:
        usage = response_json['usage']
        model = data['model']  # Get current model from request data

        print(f"Token Usage:")
        print(f"  Prompt tokens: {usage.get('prompt_tokens', 'N/A')}")
        print(f"  Completion tokens: {usage.get('completion_tokens', 'N/A')}")
        print(f"  Total tokens: {usage.get('total_tokens', 'N/A')}")
        
        # Load existing usage data if file exists
        usage_file = 'token_usage.json'
        if os.path.exists(usage_file):
            with open(usage_file, 'r') as f:
                all_models_usage = json.load(f)
        else:
            all_models_usage = {}
            
        # Initialize model usage if not exists
        if model not in all_models_usage:
            all_models_usage[model] = {
                'prompt_tokens': 0,
                'completion_tokens': 0
            }
            
        # Update usage for current model
        all_models_usage[model]['prompt_tokens'] += usage.get('prompt_tokens', 0)
        all_models_usage[model]['completion_tokens'] += usage.get('completion_tokens', 0)
        
        # Write updated usage to file
        with open(usage_file, 'w') as f:
            json.dump(all_models_usage, f, indent=2)
        
        # Add usage info to the response for API consumers
        # response_json['token_usage'] = {
        #     'prompt_tokens': usage.get('prompt_tokens', 0),
        #     'completion_tokens': usage.get('completion_tokens', 0),
        #     'total_tokens': usage.get('total_tokens', 0)
        # }
    else:
        print("No usage information available in response")
    
    return response_json

def analyze_excel_data(excel_data, question):
    """
    Analyze Excel data using LLM with structured prompt
    
    Args:
        excel_data (dict): The extracted Excel data in JSON format
        question (str): User's question about the data
    
    Returns:
        dict: Contains 'success', 'answer', 'raw_response', and 'error' fields
    """
    try:
        # Create structured prompt
        prompt = f"""
You are an expert at analyzing Excel spreadsheet data. Below is the extracted content from an Excel file in JSON format, followed by a user's question.

Excel Data (JSON format):
{json.dumps(excel_data, indent=2)}

User Question: {question}

Instructions:
- Analyze the Excel data and answer the user's question.
- If multiple insights or observations are relevant, list each separately.
- Each answer should include an "answer" field with plain text (no markdown), and if applicable, an "attribution" field listing all relevant cell coordinates.
- The response must be a valid JSON array of objects.
- Use this exact format for each item:
[
  {{
    "answer": "A detailed summary/insight about the findings in plain text"
  }},
  {{
    "answer": "First insight here in plain text",
    "attribution": ["sheet_name", "cell1", "cell2"]
  }}
]

- If attribution is not applicable, only include the "answer" field for that item.
- Do not use markdown formatting or backticks in the response.
- The output must be directly usable as JSON.
- Return only the JSON array, nothing else.

Answer:
"""
        
        # Make LLM call
        llm_response = make_llm_call(prompt)
        raw_answer = llm_response['choices'][0]['message']['content']
        
        # Try to parse the response as JSON
        try:
            parsed_answer = json.loads(raw_answer.strip())
            print(parsed_answer)
            return {
                'success': True,
                'answer': parsed_answer,
                'raw_response': raw_answer,
                'error': None
            }
        except json.JSONDecodeError:
            # If JSON parsing fails, return raw text
            print("JSON parsing failed")
            print(raw_answer)
            return {
                'success': True,
                'answer': raw_answer,
                'raw_response': raw_answer,
                'error': 'Response was not in expected JSON format'
            }
            
    except Exception as e:
        return {
            'success': False,
            'answer': None,
            'raw_response': None,
            'error': f'Error in LLM analysis: {str(e)}'
        }

def main():
    prompt = "What is the capital of France?"
    response = make_llm_call(prompt)
    print(f"\nResponse: {response['choices'][0]['message']['content']}")

if __name__ == "__main__":
    main()