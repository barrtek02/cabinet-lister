import os
import subprocess
import threading
import itertools
import time

# Specify the name of your Python script here
script_name = "Program.py"

# Check if the script exists
if not os.path.exists(script_name):
    print(f"Error: {script_name} does not exist")
else:
    # Define a spinner function
    def spinner():
        for char in itertools.cycle("|/-\\"):
            if done:
                break
            print(char, end="", flush=True)
            time.sleep(0.1)
            print("\b", end="", flush=True)

    # Start the spinner in a separate thread
    done = False
    spinner_thread = threading.Thread(target=spinner)
    spinner_thread.start()

    # Run pyinstaller to create the executable
    result = subprocess.run(
        [
            "pyinstaller",
            "--onefile",
            "--noconsole",
            "--name=Rozpiski",
            "--icon=icon.ico",
            script_name,
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )

    # Stop the spinner
    done = True
    spinner_thread.join()

    # Check if pyinstaller ran successfully
    if result.returncode != 0:
        print(f"Error: pyinstaller failed with code {result.returncode}")
        print(f"Output: {result.stdout.decode()}")
        print(f"Error: {result.stderr.decode()}")
    else:
        print("Executable created successfully in the 'dist' directory")
