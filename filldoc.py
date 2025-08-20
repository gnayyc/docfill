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


def main():
    """Main application entry point."""
    parser = argparse.ArgumentParser(
        description="Fill DOCX templates with data from configuration files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python filldoc.py template.docx config.yaml -o output.docx
  python filldoc.py template.docx config.json
  python filldoc.py template.docx config.ini --check-placeholders
        """
    )
    
    parser.add_argument(
        'template',
        type=str,
        help='Path to the DOCX template file'
    )
    
    parser.add_argument(
        'config',
        type=str,
        help='Path to the configuration file (YAML, JSON, or INI)'
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
    
    args = parser.parse_args()
    
    try:
        # Validate input files
        template_path = Path(args.template)
        config_path = Path(args.config)
        
        if not template_path.exists():
            print(f"Error: Template file not found: {args.template}", file=sys.stderr)
            return 1
        
        if not config_path.exists():
            print(f"Error: Config file not found: {args.config}", file=sys.stderr)
            return 1
        
        # Initialize processors
        if args.verbose:
            print(f"Loading template: {template_path}")
            print(f"Loading config: {config_path}")
        
        processor = DocxProcessor(template_path)
        reader = ConfigReader(config_path)
        
        # Read configuration data
        config_data = reader.read()
        
        if args.verbose:
            print(f"Loaded {len(config_data)} configuration items")
        
        # Check placeholders mode
        if args.check_placeholders:
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
        
        # Determine output path
        if args.output:
            output_path = Path(args.output)
        else:
            # Default: add "_filled" before the extension
            stem = template_path.stem
            suffix = template_path.suffix
            output_path = template_path.parent / f"{stem}_filled{suffix}"
        
        # Process the template
        if args.verbose:
            print(f"Processing template...")
        
        processor.fill_template(config_data, output_path)
        
        print(f"Template filled successfully: {output_path}")
        
        # Show summary
        if args.verbose:
            placeholders = processor.get_placeholders()
            config_keys = set(config_data.keys())
            matched = len(placeholders & config_keys)
            print(f"Placeholders processed: {matched}/{len(placeholders)}")
        
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())