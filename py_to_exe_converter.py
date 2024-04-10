import os
import subprocess

script_name = "Program.py"

if not os.path.exists(script_name):
    print(f"Error: {script_name} does not exist")
else:
    result = subprocess.run(
        ["pyinstaller", "--onefile", script_name],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )

    if result.returncode != 0:
        print(f"Error: pyinstaller failed with code {result.returncode}")
        print(f"Output: {result.stdout.decode()}")
        print(f"Error: {result.stderr.decode()}")
    else:
        print(f"Executable created successfully in the 'dist' directory")