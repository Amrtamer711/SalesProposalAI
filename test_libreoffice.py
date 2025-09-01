#!/usr/bin/env python3
"""Test LibreOffice installation and PDF conversion."""

import subprocess
import tempfile
import os
import shutil
from pathlib import Path

def test_libreoffice():
    print("=" * 60)
    print("LIBREOFFICE INSTALLATION TEST")
    print("=" * 60)
    
    # Check various LibreOffice commands
    commands = [
        'libreoffice',
        'soffice', 
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        '/opt/libreoffice/program/soffice'
    ]
    
    found_cmd = None
    for cmd in commands:
        print(f"\nChecking {cmd}...")
        if shutil.which(cmd):
            print(f"  ✓ Found in PATH")
            found_cmd = cmd
        elif os.path.exists(cmd):
            print(f"  ✓ Found at {cmd}")
            found_cmd = cmd
        else:
            print(f"  ✗ Not found")
    
    if not found_cmd:
        print("\n❌ LibreOffice not found!")
        return False
    
    # Check version
    print(f"\nChecking version of {found_cmd}...")
    try:
        result = subprocess.run([found_cmd, '--version'], 
                              capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print(f"✓ Version: {result.stdout.strip()}")
        else:
            print(f"✗ Error getting version: {result.stderr}")
    except Exception as e:
        print(f"✗ Exception: {e}")
    
    # Test conversion with a simple file
    print("\nTesting PDF conversion...")
    try:
        # Create a simple text file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
            f.write("Test document for PDF conversion")
            test_file = f.name
        
        # Try to convert to PDF
        output_dir = tempfile.gettempdir()
        cmd = [found_cmd, '--headless', '--convert-to', 'pdf', 
               '--outdir', output_dir, test_file]
        
        print(f"Running: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            pdf_path = os.path.join(output_dir, 
                                   Path(test_file).stem + '.pdf')
            if os.path.exists(pdf_path):
                print(f"✓ PDF created successfully at {pdf_path}")
                print(f"  Size: {os.path.getsize(pdf_path)} bytes")
                os.unlink(pdf_path)
            else:
                print(f"✗ PDF not found at expected location: {pdf_path}")
        else:
            print(f"✗ Conversion failed with code {result.returncode}")
            print(f"  stdout: {result.stdout}")
            print(f"  stderr: {result.stderr}")
        
        # Cleanup
        os.unlink(test_file)
        
    except Exception as e:
        print(f"✗ Test failed with exception: {e}")
    
    print("\n" + "=" * 60)
    return found_cmd is not None

if __name__ == "__main__":
    test_libreoffice()