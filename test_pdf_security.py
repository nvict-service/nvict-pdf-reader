#!/usr/bin/env python3
"""
Test script to check PDF security detection capabilities with PyMuPDF (fitz)
"""

import fitz  # PyMuPDF

def check_pdf_security(pdf_path):
    """Check various security aspects of a PDF"""
    try:
        doc = fitz.open(pdf_path)

        print(f"PDF: {pdf_path}")
        print("=" * 60)

        # Check encryption
        print(f"1. Is Encrypted: {doc.is_encrypted}")
        print(f"   Description: Is the document password-protected?")

        # Check if requires password
        print(f"\n2. Is PDF: {doc.is_pdf}")
        print(f"   Description: Is this a valid PDF?")

        # Check metadata for signatures
        metadata = doc.metadata
        print(f"\n3. Metadata:")
        for key, value in metadata.items():
            if value:
                print(f"   {key}: {value}")

        # Check for signature fields
        print(f"\n4. Looking for Signature Fields...")
        has_signatures = False
        try:
            # Try to find signature annotations
            for page_num in range(len(doc)):
                page = doc[page_num]
                annots = page.annots()
                if annots:
                    for annot in annots:
                        if annot:
                            annot_dict = annot.info
                            if annot_dict.get('subtype') == 'Sig':
                                has_signatures = True
                                print(f"   ✓ Signature field found on page {page_num + 1}")
        except Exception as e:
            print(f"   Could not check signatures: {e}")

        if not has_signatures:
            print(f"   No signature fields found")

        # Check document permissions (if encrypted)
        print(f"\n5. Document Permissions:")
        try:
            # These properties indicate what operations are allowed
            can_print = doc.permissions
            print(f"   Permissions flags: {bin(doc.permissions)}")
            # Common permission bits:
            # Bit 2 (4): Print
            # Bit 3 (8): Modify contents
            # Bit 4 (16): Copy text
            # Bit 5 (32): Add annotations
            # Bit 8 (256): Assemble
            # Bit 9 (512): Print quality

            if doc.permissions & 4:
                print(f"   ✓ Can Print")
            else:
                print(f"   ✗ Cannot Print")

            if doc.permissions & 8:
                print(f"   ✓ Can Modify")
            else:
                print(f"   ✗ Cannot Modify")

            if doc.permissions & 16:
                print(f"   ✓ Can Copy Text")
            else:
                print(f"   ✗ Cannot Copy Text")

            if doc.permissions & 32:
                print(f"   ✓ Can Add Annotations")
            else:
                print(f"   ✗ Cannot Add Annotations")
        except Exception as e:
            print(f"   Could not determine permissions: {e}")

        doc.close()
        print("\n" + "=" * 60)

    except Exception as e:
        print(f"Error opening PDF: {e}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        pdf_file = sys.argv[1]
        check_pdf_security(pdf_file)
    else:
        print("Usage: python test_pdf_security.py <pdf_file>")
        print("Example: python test_pdf_security.py document.pdf")
