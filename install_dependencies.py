import subprocess
import sys

packages = [
    "selenium",
    "webdriver-manager",
    "psutil"
]

def install(package):
    try:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"✅ Installed: {package}\n")
    except Exception as e:
        print(f"❌ Failed to install {package}: {e}")

def main():
    print("Starting installation of required libraries...\n")
    for pkg in packages:
        install(pkg)
    print("All installations completed.")

if __name__ == "__main__":
    main()
