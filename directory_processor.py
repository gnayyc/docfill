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
        """Find all DOCX files in the given directory."""
        path = Path(directory_path)
        return list(path.glob("*.docx"))

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

        # Choose iterator
        file_iter = in_path.rglob("*.docx") if recursive else in_path.glob("*.docx")

        outputs: List[Path] = []
        for template_file in file_iter:
            stem = template_file.stem
            suffix = template_file.suffix
            out_name = f"{stem}_filled{suffix}" if add_filled_suffix else f"{stem}{suffix}"
            out_file = out_path / out_name

            if verbose:
                print(f"Processing: {template_file} -> {out_file}")

            processor = DocxProcessor(template_file)
            processor.fill_template(config_data, out_file)
            outputs.append(out_file)

        if verbose:
            print(f"Completed processing {len(outputs)} file(s)")

        return outputs
