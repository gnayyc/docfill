#!/usr/bin/env python3
"""
FillDoc - DOCX Template Filler

A tool to fill DOCX templates with data from configuration files (YAML, JSON, INI).
Uses {{placeholder}} format in templates.
"""

import argparse
import sys
from pathlib import Path
from config_reader import ConfigReader
from docx_processor import DocxProcessor
from directory_processor import DirectoryProcessor
from pdf_processor import PdfProcessor


def main():
    """Main application entry point."""
    parser = argparse.ArgumentParser(
        description="Fill DOCX templates with data from configuration files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python docfill.py config.yaml template.docx -o output.docx
  python docfill.py config.json template.docx --pdf
  python docfill.py config.yaml template1.docx template2.docx --pdf
  python docfill.py config.yaml templates/ --pdf -v
  python docfill.py config.ini template.docx --check-placeholders
        """
    )
    
    parser.add_argument(
        'config',
        type=str,
        help='Path to the configuration file (YAML, JSON, or INI)'
    )

    parser.add_argument(
        'templates',
        nargs='+',
        help='Path(s) to DOCX template file(s) or directory(ies)'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=str,
        help='Output file path (default: adds "_filled" to template name)'
    )
    
    parser.add_argument(
        '--check-placeholders',
        action='store_true',
        help='Show all placeholders found in template and available config keys'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Show verbose output'
    )

    parser.add_argument(
        '--pdf',
        action='store_true',
        help='Also generate PDF output'
    )

    parser.add_argument(
        '--prefer-libre',
        action='store_true',
        help='Prefer LibreOffice for PDF conversion (avoids Word permissions but may have layout differences)'
    )
    
    args = parser.parse_args()
    
    try:
        # Validate config file
        config_path = Path(args.config)
        if not config_path.exists():
            print(f"Error: Config file not found: {args.config}", file=sys.stderr)
            return 1

        # Validate template paths
        template_paths = []
        for template_str in args.templates:
            template_path = Path(template_str)
            if not template_path.exists():
                print(f"Error: Template path not found: {template_str}", file=sys.stderr)
                return 1
            template_paths.append(template_path)

        # Load configuration data once
        reader = ConfigReader(config_path)
        config_data = reader.read()

        if args.verbose:
            print(f"Loading config: {config_path}")
            print(f"Loaded {len(config_data)} configuration items")

        # Process all template paths
        all_output_files = []
        directories_processed = []
        files_processed = []

        for template_path in template_paths:
            if template_path.is_dir():
                # Directory mode
                output_dir = Path(args.output) if args.output else (template_path / "output")
                if args.verbose:
                    print(f"Directory mode detected. Input dir: {template_path}")
                    print(f"Output dir: {output_dir}")

                dp = DirectoryProcessor()
                output_files = dp.process_directory(
                    input_dir=template_path,
                    config_path=config_path,
                    output_dir=output_dir,
                    add_filled_suffix=True,
                    recursive=False,
                    verbose=args.verbose,
                )
                all_output_files.extend(output_files)
                directories_processed.append(template_path)

            else:
                # File mode
                if args.verbose:
                    print(f"Processing file: {template_path}")

                # Determine output path
                if args.output and len(template_paths) == 1:
                    output_path = Path(args.output)
                else:
                    # Default: add "_filled" before the extension
                    stem = template_path.stem
                    suffix = template_path.suffix
                    output_path = template_path.parent / f"{stem}_filled{suffix}"

                # Check placeholders mode for single file
                if args.check_placeholders and len(template_paths) == 1:
                    processor = DocxProcessor(template_path)
                    placeholders = processor.get_placeholders()
                    config_keys = set(config_data.keys())

                    print(f"\nTemplate placeholders found ({len(placeholders)}):")
                    for placeholder in sorted(placeholders):
                        status = "✓" if placeholder in config_keys else "✗"
                        print(f"  {status} {{{{{placeholder}}}}}")

                    print(f"\nConfiguration keys available ({len(config_keys)}):")
                    for key in sorted(config_keys):
                        used = "✓" if key in placeholders else "✗"
                        print(f"  {used} {key} = {config_data[key][:50]}{'...' if len(config_data[key]) > 50 else ''}")

                    missing = placeholders - config_keys
                    unused = config_keys - placeholders

                    if missing:
                        print(f"\nMissing config values for placeholders ({len(missing)}):")
                        for placeholder in sorted(missing):
                            print(f"  ✗ {placeholder}")

                    if unused:
                        print(f"\nUnused config values ({len(unused)}):")
                        for key in sorted(unused):
                            print(f"  ✗ {key}")

                    return 0

                # Process the template
                try:
                    processor = DocxProcessor(template_path)
                    processor.fill_template(config_data, output_path, args.verbose)
                    all_output_files.append(output_path)
                    files_processed.append(template_path)
                    print(f"Template filled successfully: {output_path}")
                except Exception as e:
                    print(f"Error processing template '{template_path}': {e}", file=sys.stderr)
                    if args.verbose:
                        import traceback
                        traceback.print_exc()
                    continue

        # Generate PDF files if requested
        if args.pdf and all_output_files:
            if args.verbose:
                print("Converting DOCX files to PDF...")

            pdf_processor = PdfProcessor(prefer_libre=args.prefer_libre)
            if pdf_processor.get_available_method() == "none":
                print("Warning: No PDF conversion method available. Install LibreOffice or docx2pdf for PDF output.")
            else:
                if args.verbose:
                    print(f"Using PDF conversion method: {pdf_processor.get_available_method()}")

                for docx_file in all_output_files:
                    try:
                        pdf_file = pdf_processor.convert_to_pdf(docx_file)
                        if args.verbose:
                            print(f"PDF created: {pdf_file}")
                    except Exception as e:
                        print(f"Warning: Failed to convert {docx_file} to PDF: {e}")

        # Summary
        if directories_processed:
            for dir_path in directories_processed:
                print(f"Directory processed: {dir_path}")
        if files_processed:
            print(f"Files processed: {len(files_processed)}")

        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
