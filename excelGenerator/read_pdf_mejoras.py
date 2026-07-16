#!/usr/bin/env python3
"""
STARK096 — ALDEA: Read PDF Mejoras
Extract improvement options from PDF catalog
"""

import PyPDF2
import re
import json
from utils import print_success, print_error, print_info, print_warning


def extract_text_from_pdf(pdf_path):
    """Extract all text from PDF"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text
    except Exception as e:
        print_error(f"Error reading PDF: {e}")
        return None


def parse_mejoras_catalog(text):
    """
    Parse the mejoras catalog from PDF text
    Returns a structured dict with concepts, options, prices, and restrictions
    """
    mejoras = {
        "conceptos": [],
        "metadata": {
            "source": "6_LISTADO MEJORAS Neptuno C63.pdf",
            "parsed_date": "2026-07-17"
        }
    }
    
    # This is a template parser - will be refined after seeing the actual PDF structure
    # For now, let's just extract the text and save it for analysis
    
    return mejoras, text


def read_mejoras_pdf():
    """Main function to read and parse mejoras PDF"""
    
    pdf_path = "../dosier/6_LISTADO MEJORAS Neptuno C63.pdf"
    
    print_info("Reading mejoras catalog from PDF...")
    
    # Extract text
    text = extract_text_from_pdf(pdf_path)
    if not text:
        return 1
    
    # Parse mejoras
    mejoras, raw_text = parse_mejoras_catalog(text)
    
    # Save raw text for analysis
    with open("mejoras_raw_text.txt", "w", encoding="utf-8") as f:
        f.write(raw_text)
    print_success(f"Raw text saved to mejoras_raw_text.txt ({len(raw_text)} characters)")
    
    # Save structured data
    with open("mejoras_catalog.json", "w", encoding="utf-8") as f:
        json.dump(mejoras, f, indent=2, ensure_ascii=False)
    print_success("Catalog saved to mejoras_catalog.json")
    
    # Print preview
    print_info("=" * 60)
    print_info("PDF TEXT PREVIEW (first 2000 chars):")
    print_info("=" * 60)
    print(raw_text[:2000])
    print_info("=" * 60)
    
    return 0


if __name__ == "__main__":
    exit(read_mejoras_pdf())
