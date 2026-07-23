import os
from config import MAX_CHARS

#exported to call_functions.py
schema_get_file_content = {
    "type": "function",
    "function": {
        "name": "get_file_content",
        "description": f"Retrieves the content (at most {MAX_CHARS} characters) of a specified file within the working directory",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the file to read, relative to the working directory",
                },
            },
            "required": ["file_path"],
        },
    },
}

def get_file_content(working_directory: str, file_path: str) -> str:
    try:
        working_dir = os.path.abspath(working_directory)
        target_fpath = os.path.join(working_dir, file_path)
        target_fpath = os.path.normpath(target_fpath)
        valid_target_fpath = os.path.commonpath([working_dir, target_fpath]) == working_dir
        
        if not valid_target_fpath:
            return f"Error: Cannot read '{file_path}' as it is outside the permitted working directory"
        
        if not os.path.isfile(target_fpath):
            return f'Error: File not found or is not a regular file: "{file_path}"'
        
        with open(target_fpath, "r") as f:
            content = f.read(MAX_CHARS)
            if f.read(1):
                content += f"[... File '{file_path}' truncated at {MAX_CHARS} characters]"
            return content
        
    except Exception as e:
        return f"Error: {e}"
