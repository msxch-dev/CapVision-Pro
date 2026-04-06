import subprocess
import sys
import importlib
# CapVision Pro Dependencies Installer
# Run this file to install required libraries automatically
pkgs = [
    "opencv-python",
    "sounddevice",
    "pillow",
    "win10toast",
    "pystray",
    "psutil",
    "pywin32",
    "pyaudio",
    "wave",
    "audioop"
]

for pkg in pkgs:
    try:
        importlib.import_module(pkg.replace("-", "_"))
    except:
        subprocess.run([sys.executable, "-m", "pip", "install", pkg])

input("Press Enter to Continue...")