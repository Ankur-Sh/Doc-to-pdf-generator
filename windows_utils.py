"""
Windows-specific utilities for automatic error handling
"""
import os
import platform
import subprocess
import time
import warnings

def is_windows():
    """Check if running on Windows"""
    return platform.system() == "Windows"

def close_word_processes():
    """Automatically close Microsoft Word processes on Windows"""
    if not is_windows():
        return False
    
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/IM", "WINWORD.EXE"],
            capture_output=True,
            timeout=5,
            shell=True
        )
        if result.returncode == 0:
            time.sleep(1)  # Give it time to close
            return True
    except:
        pass
    return False

def wait_for_file_unlock(file_path, max_retries=5, delay=1):
    """Wait for a file to be unlocked, with automatic retries"""
    for attempt in range(max_retries):
        try:
            # Try to open the file for writing
            if os.path.exists(file_path):
                with open(file_path, 'r+b') as f:
                    pass
            # Try to create/delete the file
            test_path = file_path + ".test"
            with open(test_path, 'wb') as f:
                f.write(b'test')
            os.remove(test_path)
            return True
        except (PermissionError, OSError):
            if attempt < max_retries - 1:
                time.sleep(delay)
                # Try closing Word on Windows
                if is_windows() and attempt == 1:
                    close_word_processes()
                    time.sleep(1)
            else:
                return False
    return False

def safe_save_document(document, file_path, max_retries=3):
    """Safely save a document with automatic retry and error handling"""
    file_path = os.path.normpath(file_path)
    
    # Ensure directory exists
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    for attempt in range(max_retries):
        try:
            # Check and wait for file unlock
            if not wait_for_file_unlock(file_path, max_retries=3, delay=0.5):
                # If file is locked, try closing Word
                if is_windows() and attempt == 0:
                    close_word_processes()
                    time.sleep(1)
                    continue
            
            # Remove existing file if it exists (to avoid conflicts)
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    time.sleep(0.2)  # Brief pause after deletion
                except PermissionError:
                    if attempt < max_retries - 1:
                        if is_windows():
                            close_word_processes()
                        time.sleep(1)
                        continue
                    raise
            
            # Save the document
            document.save(file_path)
            
            # Verify it was saved
            if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                return True
            else:
                raise Exception("File was created but is empty")
                
        except PermissionError:
            if attempt < max_retries - 1:
                if is_windows():
                    close_word_processes()
                time.sleep(1)
            else:
                raise
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(0.5)
            else:
                raise
    
    return False

def suppress_warnings():
    """Suppress common warnings that don't affect functionality"""
    warnings.filterwarnings('ignore', category=UserWarning, module='docx.styles')

def normalize_path(path):
    """Normalize path for Windows compatibility"""
    return os.path.normpath(os.path.abspath(path))

def ensure_directory(path):
    """Ensure directory exists, create if it doesn't"""
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except Exception:
        return False






