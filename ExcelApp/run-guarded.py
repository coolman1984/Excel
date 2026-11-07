"""Run local build/test commands with Windows crash dialogs disabled from process creation."""
import ctypes
import subprocess
import sys
ctypes.windll.kernel32.SetErrorMode(0x8003)
try:
    process = subprocess.Popen(sys.argv[1:], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, creationflags=subprocess.CREATE_NO_WINDOW)
    for line in iter(process.stdout.readline, b""):
        sys.stdout.buffer.write(line)
        sys.stdout.buffer.flush()
    sys.exit(process.wait())
except Exception as error:
    print(type(error).__name__)
    sys.exit(1)
