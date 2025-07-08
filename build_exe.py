#!/usr/bin/env python3
"""
Build Script for Radiologist Report Generator
Creates a standalone executable using PyInstaller
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        print("✅ PyInstaller is already installed")
        return True
    except ImportError:
        print("❌ PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✅ PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("❌ Failed to install PyInstaller")
            return False

def clean_build_dirs():
    """Clean previous build directories"""
    dirs_to_clean = ["build", "dist", "__pycache__"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"🧹 Cleaning {dir_name}...")
            shutil.rmtree(dir_name)
    
    # Clean .spec files
    for spec_file in Path(".").glob("*.spec"):
        print(f"🧹 Removing {spec_file}...")
        spec_file.unlink()

def build_executable():
    """Build the executable using PyInstaller"""
    print("🚀 Building executable...")
    
    # PyInstaller command with options
    cmd = [
        "pyinstaller",
        "--onefile",                    # Create single executable
        "--windowed",                   # No console window
        "--name=RadiologistApp",        # Executable name
        "--icon=icon.ico",              # Icon (if exists)
        "--add-data=admin_settings.json:.",  # Include admin settings
        "--add-data=Reports:Reports",   # Include Reports folder
        "--hidden-import=tkinter",      # Ensure tkinter is included
        "--hidden-import=tkinter.ttk",
        "--hidden-import=docx",
        "--hidden-import=docx.enum.text",
        "--hidden-import=docx.shared",
        "--hidden-import=docx.oxml",
        "--hidden-import=docx.oxml.ns",
        "radiologist_app.py"
    ]
    
    # Remove icon option if icon doesn't exist
    if not os.path.exists("icon.ico"):
        cmd.remove("--icon=icon.ico")
    
    try:
        subprocess.check_call(cmd)
        print("✅ Build completed successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        return False

def create_icon():
    """Create a simple icon file if it doesn't exist"""
    if not os.path.exists("icon.ico"):
        print("📝 Creating simple icon...")
        try:
            # Create a simple icon using PIL if available
            try:
                from PIL import Image, ImageDraw, ImageFont
                
                # Create a 256x256 image
                img = Image.new('RGBA', (256, 256), (59, 130, 246, 255))  # Blue background
                draw = ImageDraw.Draw(img)
                
                # Draw a simple medical cross
                draw.rectangle([80, 120, 176, 136], fill='white')  # Horizontal
                draw.rectangle([120, 80, 136, 176], fill='white')  # Vertical
                
                # Save as ICO
                img.save("icon.ico", format='ICO')
                print("✅ Icon created successfully")
            except ImportError:
                print("⚠️  PIL not available, skipping icon creation")
        except Exception as e:
            print(f"⚠️  Could not create icon: {e}")

def main():
    """Main build process"""
    print("🏥 Radiologist Report Generator - Build Script")
    print("=" * 50)
    
    # Check if main app file exists
    if not os.path.exists("radiologist_app.py"):
        print("❌ radiologist_app.py not found!")
        return False
    
    # Check PyInstaller
    if not check_pyinstaller():
        return False
    
    # Clean previous builds
    clean_build_dirs()
    
    # Create icon
    create_icon()
    
    # Build executable
    if build_executable():
        print("\n🎉 Build completed successfully!")
        print("📁 Executable location: dist/RadiologistApp.exe")
        print("📦 Size: ~", end="")
        
        # Show file size
        exe_path = "dist/RadiologistApp.exe"
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"{size_mb:.1f} MB")
        
        print("\n📋 Next steps:")
        print("1. Copy dist/RadiologistApp.exe to your target machine")
        print("2. Copy admin_settings.json (if you have custom settings)")
        print("3. Copy Reports folder (if you have existing reports)")
        print("4. Run RadiologistApp.exe")
        
        return True
    else:
        print("❌ Build failed!")
        return False

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1) 