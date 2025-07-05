#!/usr/bin/env python3
"""
Build script for creating Windows executable of the Radiologist Report Generator
"""

import os
import sys
import subprocess
import shutil

def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        print("✓ PyInstaller is installed")
        return True
    except ImportError:
        print("✗ PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✓ PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("✗ Failed to install PyInstaller")
            return False

def build_executable():
    """Build the executable using PyInstaller"""
    print("Building Radiologist Report Generator executable...")
    
    # PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",                    # Create single executable
        "--windowed",                   # Hide console window
        "--name=RadiologistReportGenerator",  # Executable name
        "--add-data=README.md;.",       # Include README
        "--icon=NONE",                  # No icon for now
        "radiologist_app.py"
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✓ Executable built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Build failed: {e}")
        return False

def cleanup_build_files():
    """Clean up build artifacts"""
    print("Cleaning up build files...")
    
    # Remove build directory
    if os.path.exists("build"):
        shutil.rmtree("build")
        print("✓ Removed build/ directory")
    
    # Remove spec file
    if os.path.exists("RadiologistReportGenerator.spec"):
        os.remove("RadiologistReportGenerator.spec")
        print("✓ Removed .spec file")

def main():
    print("=" * 50)
    print("Radiologist Report Generator - Build Script")
    print("=" * 50)
    
    # Check if main app file exists
    if not os.path.exists("radiologist_app.py"):
        print("✗ radiologist_app.py not found in current directory")
        print("Please run this script from the project directory")
        return False
    
    # Check PyInstaller
    if not check_pyinstaller():
        return False
    
    # Build executable
    if not build_executable():
        return False
    
    # Check if executable was created
    exe_path = os.path.join("dist", "RadiologistReportGenerator.exe")
    if os.path.exists(exe_path):
        print(f"✓ Executable created: {exe_path}")
        print(f"  Size: {os.path.getsize(exe_path) / (1024*1024):.1f} MB")
        
        # Ask if user wants to clean up build files
        response = input("\nClean up build files? (y/n): ").lower().strip()
        if response in ['y', 'yes']:
            cleanup_build_files()
        
        print("\n" + "=" * 50)
        print("BUILD COMPLETE!")
        print("=" * 50)
        print(f"Your executable is ready: {exe_path}")
        print("You can now distribute this file to other Windows computers.")
        print("Note: The executable will create a Reports/ folder when first run.")
        return True
    else:
        print("✗ Executable not found after build")
        return False

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1) 