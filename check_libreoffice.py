#!/usr/bin/env python3
"""
Diagnostic script to check if LibreOffice is properly installed and accessible
Run this to troubleshoot PDF conversion issues
"""
import os
import subprocess
import platform

def check_libreoffice():
    print("=" * 60)
    print("LibreOffice Installation Checker")
    print("=" * 60)
    print()
    
    is_windows = platform.system() == "Windows"
    
    # Check common installation paths
    if is_windows:
        program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
        paths_to_check = [
            os.path.join(program_files, "LibreOffice", "program", "soffice.exe"),
            os.path.join(program_files_x86, "LibreOffice", "program", "soffice.exe"),
            "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
        ]
    else:
        paths_to_check = [
            "/opt/homebrew/bin/soffice",
            "/usr/local/bin/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
    
    print("Checking for LibreOffice installation...")
    print()
    
    found = False
    for path in paths_to_check:
        if os.path.exists(path):
            print(f"✓ Found LibreOffice at: {path}")
            found = True
            
            # Try to run it
            try:
                if is_windows:
                    result = subprocess.run([path, "--version"], 
                                           capture_output=True, timeout=5, shell=True, text=True)
                else:
                    result = subprocess.run([path, "--version"], 
                                           capture_output=True, timeout=2, text=True)
                
                if result.returncode == 0:
                    print(f"✓ LibreOffice is working!")
                    if result.stdout:
                        print(f"  Version info: {result.stdout.strip()[:100]}")
                else:
                    print(f"✗ LibreOffice found but returned error code: {result.returncode}")
            except Exception as e:
                print(f"✗ Error running LibreOffice: {e}")
            
            break
    
    # Check if soffice is in PATH
    print()
    print("Checking if 'soffice' is in PATH...")
    try:
        if is_windows:
            result = subprocess.run(["soffice.exe", "--version"], 
                                   capture_output=True, timeout=5, shell=True, text=True)
        else:
            result = subprocess.run(["soffice", "--version"], 
                                   capture_output=True, timeout=2, text=True)
        
        if result.returncode == 0:
            print("✓ 'soffice' is available in PATH")
            if result.stdout:
                print(f"  Version: {result.stdout.strip()[:100]}")
        else:
            print("✗ 'soffice' is in PATH but not working")
    except FileNotFoundError:
        print("✗ 'soffice' is NOT in PATH")
    except Exception as e:
        print(f"✗ Error checking PATH: {e}")
    
    print()
    print("=" * 60)
    
    if not found:
        print("❌ LibreOffice NOT FOUND")
        print()
        print("To install LibreOffice:")
        if is_windows:
            print("1. Visit: https://www.libreoffice.org/download/")
            print("2. Download and install LibreOffice")
            print("3. Restart Command Prompt after installation")
            print("4. Run this script again to verify")
        else:
            print("  macOS: brew install --cask libreoffice")
            print("  Linux: sudo apt install libreoffice")
    else:
        print("✓ LibreOffice is installed and should work for PDF conversion")
    
    print("=" * 60)

if __name__ == "__main__":
    check_libreoffice()







