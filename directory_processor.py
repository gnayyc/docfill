"""Directory processing functionality for DocFill."""

from pathlib import Path
from typing import Union, List, Tuple


class DirectoryProcessor:
    """Process directories containing DOCX template files."""
    
    def validate_input_directory(self, directory_path: Union[str, Path]) -> bool:
        """Validate that input directory exists."""
        path = Path(directory_path)
        return path.exists() and path.is_dir()
    
    def validate_input_directory_strict(self, directory_path: Union[str, Path]) -> None:
        """Validate that input directory exists, raise error if not."""
        path = Path(directory_path)
        if not path.exists():
            raise FileNotFoundError(f"Directory not found: {path}")
        if not path.is_dir():
            raise FileNotFoundError(f"Path is not a directory: {path}")
    
    def validate_output_directory(self, directory_path: Union[str, Path]) -> None:
        """Validate output directory, create if it doesn't exist."""
        path = Path(directory_path)
        if not path.exists():
            path.mkdir(parents=True, exist_ok=True)
    
    def find_docx_files(self, directory_path: Union[str, Path]) -> List[Path]:
        """Find all DOCX files in the given directory, excluding temporary and system files.

        Returns files sorted by filename to ensure consistent processing order.
        """
        path = Path(directory_path)
        files = list(path.glob("*.docx"))

        # Filter out temporary and system files
        filtered_files = []
        for file in files:
            filename = file.name
            # Skip temporary files (starts with ~$, ~WRL, or contains temp patterns)
            if (filename.startswith('~$') or
                filename.startswith('~WRL') or
                filename.startswith('.') or
                'temp' in filename.lower() or
                'tmp' in filename.lower() or
                filename.endswith('.tmp.docx') or
                filename.endswith('_filled.docx')):
                continue
            filtered_files.append(file)

        # Sort by filename to ensure consistent processing order
        filtered_files.sort(key=lambda x: x.name)
        return filtered_files

    def process_directory(
        self,
        input_dir: Union[str, Path],
        config_path: Union[str, Path],
        output_dir: Union[str, Path],
        *,
        add_filled_suffix: bool = True,
        recursive: bool = False,
        verbose: bool = False,
    ) -> List[Path]:
        """Fill all DOCX templates in a directory and write outputs.

        - Reads config once and applies to every template.
        - Writes outputs to `output_dir`, preserving filenames.
        - Appends "_filled" before extension by default.

        Returns a list of generated output file paths.
        """

        # Import here to avoid hard dependency at module import time in tests
        from config_reader import ConfigReader
        from docx_processor import DocxProcessor

        in_path = Path(input_dir)
        out_path = Path(output_dir)

        # Validate directories
        self.validate_input_directory_strict(in_path)
        self.validate_output_directory(out_path)

        # Load configuration once
        reader = ConfigReader(config_path)
        config_data = reader.read()

        # Choose iterator - use find_docx_files to filter out temp files
        if recursive:
            file_iter = []
            # Include root directory
            file_iter.extend(self.find_docx_files(in_path))
            # Include subdirectories in sorted order
            subdirs = [subdir for subdir in in_path.rglob("*") if subdir.is_dir() and subdir != in_path]
            subdirs.sort(key=lambda x: str(x))
            for subdir in subdirs:
                file_iter.extend(self.find_docx_files(subdir))
        else:
            file_iter = self.find_docx_files(in_path)

        total_files = len(file_iter)
        if total_files == 0:
            print("No DOCX files found to process.")
            return []

        print(f"Found {total_files} DOCX file(s) to process")

        outputs: List[Path] = []
        for i, template_file in enumerate(file_iter, 1):
            stem = template_file.stem
            suffix = template_file.suffix
            out_name = f"{stem}_filled{suffix}" if add_filled_suffix else f"{stem}{suffix}"
            out_file = out_path / out_name

            print(f"[{i}/{total_files}] Processing: {template_file.name}")
            if verbose:
                print(f"  Input:  {template_file}")
                print(f"  Output: {out_file}")

            try:
                processor = DocxProcessor(template_file)
                processor.fill_template(config_data, out_file, verbose)
                outputs.append(out_file)
                print(f"[{i}/{total_files}] ✓ Completed: {template_file.name}")
            except Exception as e:
                print(f"[{i}/{total_files}] ✗ Failed: {template_file.name}")
                print(f"  Error: {e}")
                if verbose:
                    import traceback
                    traceback.print_exc()
                # Continue processing other files instead of stopping
                continue

        if verbose:
            print(f"Completed processing {len(outputs)} file(s)")

        return outputs
