import os
import subprocess

name = "run_python_file"
description = "Executes a python file in a specified directory relative to the working directory"
working_directory_description = "The working directory in which a python file may be executed"
file_path_description = "The path of the python file to be executed, relative to the working directory"
args_description = "Optional list of extra arguments which should be taken in as input if present"

#exported to call_functions.py
schema_run_python_file = {
    "type": "function",
    "function": {
        "name": "run_python_file",
        "description": "Executes a specified Python file within the working directory and returns its output",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the Python file to run, relative to the working directory",
                },
                "args": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Optional list of arguments to pass to the Python script",
                },
            },
            "required": ["file_path"],
        },
    },
}

def run_python_file(
    working_directory: str, file_path: str, args: list[str] | None = None
) -> str:
    try:
        working_dir = os.path.abspath(working_directory)
        target_fpath = os.path.join(working_dir, file_path)
        target_fpath = os.path.normpath(target_fpath)
        valid_target_fpath = os.path.commonpath([working_dir, target_fpath]) == working_dir
        
        if not valid_target_fpath:
            return f'Error: Cannot execute "{file_path}" as it is outside the permitted working directory'
        
        if not os.path.isfile(target_fpath):
            return f'Error: "{file_path}" does not exist or is not a regular file'
        
        if not file_path.endswith('.py'):
            return f'Error: "{file_path}" is not a Python file'

        command = ["python", target_fpath]
        if args:
            command.extend(args)
        command_call: str = subprocess.run(command, capture_output=True, text=True, timeout=30)
        output: str = ""
        if command_call.returncode != 0:
            output += f"Process exited with code {command_call.returncode}"
        if command_call.stdout is None and command_call.stderr is None:
            output += "No output produced"
        else:
            output += f"STDOUT: {command_call.stdout}\nSTDERR: {command_call.stderr}"
            return output
        
    except Exception as e:
        return f"Error: executing Python file: {e}"
