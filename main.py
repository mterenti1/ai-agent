import os
import sys
import json
from openai import OpenAI
from dotenv import load_dotenv
import argparse
from prompts import system_prompt
from call_functions import call_function, available_functions

load_dotenv()
api_key = os.environ.get("OPENROUTER_API_KEY")

if api_key is None:
    raise RuntimeError("Could not fetch API key")

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=api_key,
)

model = 'openrouter/free'

parser = argparse.ArgumentParser(description="Chatbot")
parser.add_argument("user_prompt", type=str, help="User Prompt")
parser.add_argument("--verbose", action="store_true", help="Enable verbose output")
args = parser.parse_args()

messages = [
    {"role": "system", "content": system_prompt},
    {"role": "user", "content": args.user_prompt}
]

def main():
    for _ in range(20):
        response = client.chat.completions.create(
            model=model,
            messages=messages,
            tools=available_functions
        )
        if response.usage is None:
            raise RuntimeError("Failed API request")
        
        if args.verbose:
            print(f"User prompt: {args.user_prompt}")
            print(f"Prompt tokens: {response.usage.prompt_tokens}")
            print(f"Response tokens: {response.usage.completion_tokens}")
        message = response.choices[0].message
        messages.append(message)
        if not message.tool_calls:
            print(message.content)
            break
        else:
            for tool_call in message.tool_calls:
                function_args = json.loads(tool_call.function.arguments or {})
                print(function_args)
                result_message = call_function(tool_call)
                messages.append(result_message)
                if result_message["content"] is None:
                    raise ValueError(f"No content returned from {tool_call.function.name}")
                if args.verbose:
                    print(f"-> {result_message['content']}")
    if not messages:
        print(f"Error: no responses recorded")
        sys.exit(1)         



if __name__ == "__main__":
    main()
