#!/usr/bin/env python3
"""
Test suite for the PDF to PowerPoint converter.

This module provides comprehensive testing for the converter functionality,
including unit tests, integration tests, and error handling verification.
"""

import unittest
import tempfile
import os
from pathlib import Path
from unittest.mock import patch, MagicMock
import logging

logging.disable(logging.CRITICAL)

try:
    from convert_robust import PDFToPowerPointConverter
except ImportError:
    PDFToPowerPointConverter = None


class TestPDFToPowerPointConverter(unittest.TestCase):
    """Test cases for the PDFToPowerPointConverter class."""
    
    def setUp(self):
        """Set up test fixtures."""
        if PDFToPowerPointConverter is None:
            self.skipTest("PDFToPowerPointConverter not available")
        
        self.converter = PDFToPowerPointConverter()
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_init_default_values(self):
        """Test converter initialization with default values."""
        converter = PDFToPowerPointConverter()
        self.assertEqual(converter.dpi, 300)
        self.assertEqual(converter.image_format, 'png')
        self.assertEqual(converter.thread_count, 2)
        self.assertEqual(converter.progress_interval, 10)
    
    def test_init_custom_values(self):
        """Test converter initialization with custom values."""
        converter = PDFToPowerPointConverter(
            dpi=600, image_format='JPEG', thread_count=4, progress_interval=5
        )
        self.assertEqual(converter.dpi, 600)
        self.assertEqual(converter.image_format, 'jpeg')  # Should be lowercase
        self.assertEqual(converter.thread_count, 4)
        self.assertEqual(converter.progress_interval, 5)
    
    def test_init_invalid_format(self):
        """Test converter initialization with invalid image format."""
        with self.assertRaises(ValueError):
            PDFToPowerPointConverter(image_format='invalid')
    
    def test_validate_input_file_not_exists(self):
        """Test validation of non-existent input file."""
        with self.assertRaises(FileNotFoundError):
            self.converter.validate_input_file('/nonexistent/file.pdf')
    
    def test_validate_input_file_not_pdf(self):
        """Test validation of non-PDF file."""
        temp_file = Path(self.temp_dir) / 'test.txt'
        temp_file.write_text('not a pdf')
        
        with self.assertRaises(ValueError):
            self.converter.validate_input_file(str(temp_file))
    
    def test_validate_input_file_directory(self):
        """Test validation when path is a directory."""
        with self.assertRaises(ValueError):
            self.converter.validate_input_file(self.temp_dir)
    
    def test_validate_output_path_auto_extension(self):
        """Test output path validation with automatic extension."""
        output_path = Path(self.temp_dir) / 'test'
        result = self.converter.validate_output_path(str(output_path))
        self.assertEqual(result.suffix, '.pptx')
    
    def test_validate_output_path_creates_directory(self):
        """Test output path validation creates missing directories."""
        output_path = Path(self.temp_dir) / 'subdir' / 'test.pptx'
        result = self.converter.validate_output_path(str(output_path))
        self.assertTrue(result.parent.exists())
    
    @patch('convert_robust.convert_from_path')
    def test_convert_pdf_to_images_success(self, mock_convert):
        """Test successful PDF to images conversion."""
        mock_images = [MagicMock(), MagicMock()]
        mock_convert.return_value = mock_images
        
        pdf_file = Path(self.temp_dir) / 'test.pdf'
        pdf_file.write_bytes(b'fake pdf content')
        
        result = self.converter.convert_pdf_to_images(pdf_file)
        self.assertEqual(len(result), 2)
        mock_convert.assert_called_once()
    
    @patch('convert_robust.convert_from_path')
    def test_convert_pdf_to_images_empty_result(self, mock_convert):
        """Test PDF conversion that returns no images."""
        mock_convert.return_value = []
        
        pdf_file = Path(self.temp_dir) / 'test.pdf'
        pdf_file.write_bytes(b'fake pdf content')
        
        with self.assertRaises(RuntimeError):
            self.converter.convert_pdf_to_images(pdf_file)
    
    @patch('convert_robust.convert_from_path')
    def test_convert_pdf_to_images_exception(self, mock_convert):
        """Test PDF conversion that raises an exception."""
        mock_convert.side_effect = Exception("Conversion failed")
        
        pdf_file = Path(self.temp_dir) / 'test.pdf'
        pdf_file.write_bytes(b'fake pdf content')
        
        with self.assertRaises(RuntimeError):
            self.converter.convert_pdf_to_images(pdf_file)


class TestConverterIntegration(unittest.TestCase):
    """Integration tests for the converter (require dependencies)."""
    
    def setUp(self):
        """Set up integration test fixtures."""
        if PDFToPowerPointConverter is None:
            self.skipTest("Dependencies not available for integration tests")
        
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """Clean up integration test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_create_sample_pdf(self):
        """Create a sample PDF for testing (if reportlab is available)."""
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            
            pdf_path = Path(self.temp_dir) / 'sample.pdf'
            c = canvas.Canvas(str(pdf_path), pagesize=letter)
            c.drawString(100, 750, "Test PDF Page 1")
            c.showPage()
            c.drawString(100, 750, "Test PDF Page 2")
            c.save()
            
            self.assertTrue(pdf_path.exists())
            return pdf_path
            
        except ImportError:
            self.skipTest("reportlab not available for PDF creation")
    
    def test_end_to_end_conversion(self):
        """Test complete end-to-end conversion process."""
        try:
            import pdf2image
            import pptx
        except ImportError:
            self.skipTest("Required dependencies not available")
        
        pdf_path = self.test_create_sample_pdf()
        
        converter = PDFToPowerPointConverter(dpi=150)  # Lower DPI for faster testing
        output_path = converter.convert(str(pdf_path))
        
        self.assertTrue(output_path.exists())
        self.assertEqual(output_path.suffix, '.pptx')
        self.assertGreater(output_path.stat().st_size, 0)


class TestErrorHandling(unittest.TestCase):
    """Test error handling and edge cases."""
    
    def test_missing_dependencies_graceful_failure(self):
        """Test graceful handling when dependencies are missing."""
        pass
    
    def test_corrupted_pdf_handling(self):
        """Test handling of corrupted PDF files."""
        if PDFToPowerPointConverter is None:
            self.skipTest("PDFToPowerPointConverter not available")
        
        temp_dir = tempfile.mkdtemp()
        try:
            fake_pdf = Path(temp_dir) / 'corrupted.pdf'
            fake_pdf.write_bytes(b'This is not a valid PDF file')
            
            converter = PDFToPowerPointConverter()
            
            with self.assertRaises(RuntimeError):
                converter.convert_pdf_to_images(fake_pdf)
                
        finally:
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)


def run_tests():
    """Run all tests and return results."""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    suite.addTests(loader.loadTestsFromTestCase(TestPDFToPowerPointConverter))
    suite.addTests(loader.loadTestsFromTestCase(TestConverterIntegration))
    suite.addTests(loader.loadTestsFromTestCase(TestErrorHandling))
    
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result.wasSuccessful()


if __name__ == '__main__':
    success = run_tests()
    exit(0 if success else 1)
