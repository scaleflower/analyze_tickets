#!/usr/bin/env python3
"""
OTRS Ticket Analysis Installer
Cross-platform installer that downloads code from GitHub and sets up the environment
"""

import os
import sys
import platform
import subprocess
import urllib.request
import zipfile
import tempfile
import shutil
from pathlib import Path

def is_windows():
    """Check if running on Windows"""
    return platform.system().lower() == 'windows'

def is_linux():
    """Check if running on Linux"""
    return platform.system().lower() == 'linux'

def check_python():
    """Check if Python is installed and meets version requirements"""
    try:
        result = subprocess.run([sys.executable, '--version'], 
                              capture_output=True, text=True, check=True)
        version_output = result.stdout.strip()
        print(f"✓ Python found: {version_output}")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("✗ Python not found or not in PATH")
        return False

def check_pip():
    """Check if pip is available"""
    try:
        result = subprocess.run([sys.executable, '-m', 'pip', '--version'],
                              capture_output=True, text=True, check=True)
        pip_version = result.stdout.split()[1]
        print(f"✓ pip found: version {pip_version}")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("✗ pip not available")
        return False

def install_requirements():
    """Install required Python packages"""
    requirements = ['pandas', 'openpyxl', 'numpy']
    print("Installing required packages...")
    
    for package in requirements:
        try:
            print(f"Installing {package}...")
            subprocess.run([sys.executable, '-m', 'pip', 'install', package],
                         check=True, capture_output=True)
            print(f"✓ {package} installed successfully")
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install {package}: {e}")
            return False
    return True

def download_from_github():
    """Download the latest code from GitHub"""
    repo_url = "https://github.com/scaleflower/analyze_tickets/archive/refs/heads/master.zip"
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, "analyze_tickets.zip")
    
    print("Downloading latest code from GitHub...")
    try:
        # Download the zip file
        urllib.request.urlretrieve(repo_url, zip_path)
        print("✓ Code downloaded successfully")
        
        # Extract the zip file
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Find the extracted directory
        extracted_dir = os.path.join(temp_dir, "analyze_tickets-master")
        if not os.path.exists(extracted_dir):
            print("✗ Failed to extract code")
            return None
        
        return extracted_dir
    except Exception as e:
        print(f"✗ Download failed: {e}")
        return None

def setup_environment(target_dir):
    """Set up the environment in the target directory"""
    # Create target directory if it doesn't exist
    os.makedirs(target_dir, exist_ok=True)
    
    # Download code from GitHub
    extracted_dir = download_from_github()
    if not extracted_dir:
        return False
    
    try:
        # Copy all files to target directory
        for item in os.listdir(extracted_dir):
            source_path = os.path.join(extracted_dir, item)
            target_path = os.path.join(target_dir, item)
            
            if os.path.isdir(source_path):
                if os.path.exists(target_path):
                    shutil.rmtree(target_path)
                shutil.copytree(source_path, target_path)
            else:
                shutil.copy2(source_path, target_path)
        
        print(f"✓ Code setup complete in: {target_dir}")
        return True
    except Exception as e:
        print(f"✗ Setup failed: {e}")
        return False
    finally:
        # Clean up temporary files
        if extracted_dir:
            shutil.rmtree(os.path.dirname(extracted_dir))

def create_run_script(target_dir, is_windows):
    """Create appropriate run script for the platform"""
    if is_windows:
        script_content = '''@echo off
echo ========================================
echo OTRS Ticket Analysis
echo ========================================
echo.

python analyze_tickets.py %*

echo.
pause
'''
        script_path = os.path.join(target_dir, "run_analysis.bat")
        with open(script_path, 'w') as f:
            f.write(script_content)
        os.chmod(script_path, 0o755)
    else:
        script_content = '''#!/bin/bash
echo "========================================"
echo "OTRS Ticket Analysis"
echo "========================================"
echo

python3 analyze_tickets.py "$@"

echo
read -p "Press enter to continue..."
'''
        script_path = os.path.join(target_dir, "run_analysis.sh")
        with open(script_path, 'w') as f:
            f.write(script_content)
        os.chmod(script_path, 0o755)
    
    print(f"✓ Created run script: {script_path}")

def main():
    print("=" * 60)
    print("OTRS Ticket Analysis - Cross-Platform Installer")
    print("=" * 60)
    print()
    
    # Detect platform
    if is_windows():
        print("✓ Platform: Windows")
    elif is_linux():
        print("✓ Platform: Linux")
    else:
        print("⚠ Unsupported platform: {}".format(platform.system()))
        print("This installer supports Windows and Linux only.")
        return 1
    
    # Check Python
    if not check_python():
        print("\nPlease install Python 3.6+ from https://python.org")
        return 1
    
    # Check pip
    if not check_pip():
        print("\nPlease ensure pip is installed with your Python installation")
        return 1
    
    # Install requirements
    if not install_requirements():
        print("\nFailed to install required packages")
        return 1
    
    # Setup environment
    target_dir = os.path.join(os.getcwd(), "otrs-analysis")
    print(f"\nSetting up environment in: {target_dir}")
    
    if not setup_environment(target_dir):
        print("Failed to setup environment")
        return 1
    
    # Create run script
    create_run_script(target_dir, is_windows())
    
    print("\n" + "=" * 60)
    print("Installation Completed Successfully!")
    print("=" * 60)
    print("\nNext steps:")
    print(f"1. Change to directory: cd {target_dir}")
    
    if is_windows():
        print("2. Run analysis: run_analysis.bat [excel_file]")
        print("   or drag Excel file onto run_analysis.bat")
    else:
        print("2. Run analysis: ./run_analysis.sh [excel_file]")
        print("   or: python analyze_tickets.py [excel_file]")
    
    print("\nFor more information, see README.md")
    return 0

if __name__ == "__main__":
    sys.exit(main())
