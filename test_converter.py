#!/usr/bin/env python3
"""
Test script for PDF to PowerPoint converter

This script runs various tests to validate the converter functionality.
"""

import os
import sys
import subprocess
import tempfile
from pathlib import Path

def run_command(cmd, expect_success=True):
    """Run a command and return the result."""
    print(f"Running: {cmd}")
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    
    if expect_success and result.returncode != 0:
        print(f"❌ Command failed: {cmd}")
        print(f"Error: {result.stderr}")
        return False
    elif not expect_success and result.returncode == 0:
        print(f"❌ Command should have failed but succeeded: {cmd}")
        return False
    else:
        print(f"✅ Command {'succeeded' if expect_success else 'failed as expected'}")
        if result.stdout:
            print(f"Output: {result.stdout.strip()}")
        return True

def create_test_pdf():
    """Create a test PDF file."""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        
        pdf_path = "test_document.pdf"
        c = canvas.Canvas(pdf_path, pagesize=letter)
        
        # Page 1
        c.drawString(100, 750, 'Test Document - Page 1')
        c.drawString(100, 700, 'This is a comprehensive test of the PDF to PowerPoint converter')
        c.drawString(100, 650, 'Testing various features and error conditions')
        c.showPage()
        
        # Page 2
        c.drawString(100, 750, 'Test Document - Page 2')
        c.drawString(100, 700, 'Second page with different content')
        c.drawString(100, 650, 'Verifying multi-page conversion')
        c.showPage()
        
        # Page 3
        c.drawString(100, 750, 'Test Document - Page 3')
        c.drawString(100, 700, 'Final page of the test document')
        c.showPage()
        
        c.save()
        print(f"✅ Created test PDF: {pdf_path}")
        return pdf_path
    except ImportError:
        print("❌ reportlab not available, cannot create test PDF")
        return None

def test_basic_functionality():
    """Test basic conversion functionality."""
    print("\n=== Testing Basic Functionality ===")
    
    # Create test PDF
    pdf_path = create_test_pdf()
    if not pdf_path:
        return False
    
    # Test basic conversion
    success = run_command(f"python convert_improved.py {pdf_path} --force")
    
    # Check if output file was created
    output_path = Path(pdf_path).with_suffix('.pptx')
    if output_path.exists():
        print(f"✅ Output file created: {output_path}")
        file_size = output_path.stat().st_size / 1024  # KB
        print(f"   File size: {file_size:.1f} KB")
    else:
        print("❌ Output file not created")
        success = False
    
    # Cleanup
    if Path(pdf_path).exists():
        Path(pdf_path).unlink()
    if output_path.exists():
        output_path.unlink()
    
    return success

def test_command_line_options():
    """Test various command line options."""
    print("\n=== Testing Command Line Options ===")
    
    pdf_path = create_test_pdf()
    if not pdf_path:
        return False
    
    tests = [
        # Test help
        ("python convert_improved.py --help", True),
        
        # Test version
        ("python convert_improved.py --version", True),
        
        # Test custom output
        (f"python convert_improved.py {pdf_path} -o custom_output.pptx --force", True),
        
        # Test different DPI
        (f"python convert_improved.py {pdf_path} --dpi 150 --force", True),
        
        # Test JPEG format
        (f"python convert_improved.py {pdf_path} --format JPEG --force", True),
        
        # Test verbose mode
        (f"python convert_improved.py {pdf_path} --verbose --force", True),
    ]
    
    success = True
    for cmd, expect_success in tests:
        if not run_command(cmd, expect_success):
            success = False
    
    # Cleanup
    cleanup_files = [pdf_path, "custom_output.pptx", f"{Path(pdf_path).stem}.pptx"]
    for file in cleanup_files:
        if Path(file).exists():
            Path(file).unlink()
    
    return success

def test_error_handling():
    """Test error handling scenarios."""
    print("\n=== Testing Error Handling ===")
    
    tests = [
        # Test non-existent file
        ("python convert_improved.py nonexistent.pdf", False),
        
        # Test non-PDF file
        ("echo 'test' > test.txt && python convert_improved.py test.txt", False),
        
        # Test invalid DPI
        ("python convert_improved.py --dpi -1 test.pdf", False),
    ]
    
    success = True
    for cmd, expect_success in tests:
        if not run_command(cmd, expect_success):
            success = False
    
    # Cleanup
    if Path("test.txt").exists():
        Path("test.txt").unlink()
    
    return success

def test_original_script():
    """Test that the original script still works."""
    print("\n=== Testing Original Script Compatibility ===")
    
    pdf_path = create_test_pdf()
    if not pdf_path:
        return False
    
    # Test original script
    success = run_command(f"python convert.py {pdf_path}")
    
    # Check if output file was created
    output_path = Path(pdf_path).with_suffix('.pptx')
    if output_path.exists():
        print(f"✅ Original script output file created: {output_path}")
    else:
        print("❌ Original script output file not created")
        success = False
    
    # Cleanup
    if Path(pdf_path).exists():
        Path(pdf_path).unlink()
    if output_path.exists():
        output_path.unlink()
    
    return success

def main():
    """Run all tests."""
    print("🧪 PDF to PowerPoint Converter Test Suite")
    print("=" * 50)
    
    tests = [
        ("Basic Functionality", test_basic_functionality),
        ("Command Line Options", test_command_line_options),
        ("Error Handling", test_error_handling),
        ("Original Script Compatibility", test_original_script),
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ Test '{test_name}' failed with exception: {e}")
            results.append((test_name, False))
    
    # Summary
    print("\n" + "=" * 50)
    print("📊 Test Results Summary")
    print("=" * 50)
    
    passed = 0
    total = len(results)
    
    for test_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status} {test_name}")
        if result:
            passed += 1
    
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! The converter is working correctly.")
        return 0
    else:
        print("⚠️  Some tests failed. Please check the output above.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
