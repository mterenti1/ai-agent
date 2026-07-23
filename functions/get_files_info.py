import os

file_path_description = "Directory path to list files info from, relative to the working directory (default is the working directory itself)"

#exported to call_functions.py
schema_get_files_info = {
    "type": "function",
    "function": {
        "name": "get_files_info",
        "description": "Lists files in a specified directory relative to the working directory, providing file size and directory status",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Directory path to list files info from, relative to the working directory (default is the working directory itself)",
                }
            },
            "required": ["file_path"]
        },
    },
}

def get_files_info(working_directory: str, file_path: str = ".") -> str:
    try:
        working_dir = os.path.abspath(working_directory)
        target_dir = os.path.join(working_dir, file_path)
        target_dir = os.path.normpath(target_dir)
        valid_target_dir = os.path.commonpath([working_dir, target_dir]) == working_dir
    
        if not valid_target_dir:
            return f'Error: Cannot list "{file_path}" as it is outside the working directory "{working_directory}"'
    
        if not os.path.isdir(target_dir):
            return f'Error: "{file_path}" is not a directory'
        
        files_info: list[str] = []
        for filename in os.listdir(target_dir):
            fpath = os.path.join(target_dir, filename)
            file_size = os.path.getsize(fpath)
            is_dir = os.path.isdir(fpath)
            files_info.append(f"- {filename}: file_size={file_size}, is_dir={is_dir}")
        return "\n".join(files_info)
    except Exception as e:
        return f'Error listing files: {e}'
    
    
    
