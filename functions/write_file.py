import os
from define_schema import define_schema, FUNCTOOL_SCHEMA, Parameter

#exported to call_functions.py
schema_write_file = {
    "type": "function",
    "function": {
        "name": "write_file",
        "description": "Writes text content to a specified file within the working directory (overwriting if the file exists)",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the file to write, relative to the working directory",
                },
                "content": {
                    "type": "string",
                    "description": "Text content to write to the file",
                },
            },
            "required": ["file_path", "content"],
        },
    },
}

def write_file(working_directory: str, file_path: str, content: str) -> str:
    try:
        working_dir = os.path.abspath(working_directory)
        target_fpath = os.path.join(working_dir, file_path)
        target_fpath = os.path.normpath(target_fpath)
        valid_target_fpath = os.path.commonpath([working_dir, target_fpath]) == working_dir
        
        if not valid_target_fpath:
            return f'Error: Cannot write to "{file_path}" as it is outside the permitted working directory'
        
        if os.path.isdir(target_fpath):
            return f'Error: Cannot write to "{file_path}" as it is a directory'
        
        os.makedirs(os.path.dirname(target_fpath), exist_ok=True)

        with open(target_fpath, "w") as f:
            f.write(content)
            return f'Successfully wrote to "{file_path}" ({len(content)} characters written)'
    except Exception as e:
        return f'Error: {e}'
