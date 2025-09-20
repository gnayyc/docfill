"""PDF processor for converting filled DOCX files to PDF."""

import platform
import subprocess
from pathlib import Path
from typing import Union


class PdfProcessor:
    """Convert DOCX files to PDF using available tools."""

    def __init__(self, prefer_libre=False):
        """Initialize PDF processor and detect available conversion methods.

        Args:
            prefer_libre: If True, prefer LibreOffice over Word (avoids permissions but may have layout differences)
        """
        self.prefer_libre = prefer_libre
        self.conversion_method = self._detect_conversion_method()

    def _detect_conversion_method(self) -> str:
        """Detect which PDF conversion method is available."""
        if self.prefer_libre:
            # User explicitly wants LibreOffice first (no permissions needed)
            if self._check_libreoffice():
                return "libreoffice"

            # Try pandoc as backup (no permissions needed)
            if self._check_pandoc():
                return "pandoc"

            # Fallback to Word
            if platform.system() == "Windows":
                try:
                    import win32com.client
                    return "word_com"
                except ImportError:
                    pass

            try:
                import docx2pdf
                return "docx2pdf"
            except ImportError:
                pass
        else:
            # Default: Try Word first for best layout fidelity
            # Try Word COM on Windows first
            if platform.system() == "Windows":
                try:
                    import win32com.client
                    return "word_com"
                except ImportError:
                    pass

            # Try docx2pdf (may require Word permissions on macOS)
            try:
                import docx2pdf
                return "docx2pdf"
            except ImportError:
                pass

            # Try LibreOffice (no permissions needed)
            if self._check_libreoffice():
                return "libreoffice"

            # Try pandoc as backup (no permissions needed, good quality)
            if self._check_pandoc():
                return "pandoc"

        return "none"

    def _check_libreoffice(self) -> bool:
        """Check if LibreOffice is available."""
        try:
            result = subprocess.run(
                ["libreoffice", "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            return result.returncode == 0
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return False

    def _check_pandoc(self) -> bool:
        """Check if Pandoc is available."""
        try:
            result = subprocess.run(
                ["pandoc", "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            return result.returncode == 0
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return False

    def convert_to_pdf(self, docx_path: Union[str, Path], pdf_path: Union[str, Path] = None) -> Path:
        """Convert DOCX file to PDF."""
        docx_path = Path(docx_path)

        if not docx_path.exists():
            raise FileNotFoundError(f"DOCX file not found: {docx_path}")

        if pdf_path is None:
            pdf_path = docx_path.with_suffix('.pdf')
        else:
            pdf_path = Path(pdf_path)

        if self.conversion_method == "libreoffice":
            return self._convert_with_libreoffice(docx_path, pdf_path)
        elif self.conversion_method == "word_com":
            return self._convert_with_word_com(docx_path, pdf_path)
        elif self.conversion_method == "docx2pdf":
            return self._convert_with_docx2pdf(docx_path, pdf_path)
        elif self.conversion_method == "pandoc":
            return self._convert_with_pandoc(docx_path, pdf_path)
        else:
            raise RuntimeError("No PDF conversion method available. Install LibreOffice, Pandoc, python-docx2pdf, or use Windows with Word.")

    def _convert_with_libreoffice(self, docx_path: Path, pdf_path: Path) -> Path:
        """Convert using LibreOffice."""
        output_dir = pdf_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            result = subprocess.run([
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_dir),
                str(docx_path)
            ],
            capture_output=True,
            text=True,
            timeout=60
            )

            if result.returncode != 0:
                raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

            # LibreOffice creates PDF with same name as DOCX
            generated_pdf = output_dir / f"{docx_path.stem}.pdf"

            # Rename to desired output path if different
            if generated_pdf != pdf_path:
                generated_pdf.rename(pdf_path)

            return pdf_path

        except subprocess.TimeoutExpired:
            raise RuntimeError("LibreOffice conversion timed out")

    def _convert_with_word_com(self, docx_path: Path, pdf_path: Path) -> Path:
        """Convert using Word COM (Windows only)."""
        import win32com.client

        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False

            doc = word.Documents.Open(str(docx_path.absolute()))
            doc.SaveAs2(str(pdf_path.absolute()), FileFormat=17)  # 17 = PDF format
            doc.Close()

            return pdf_path

        except Exception as e:
            raise RuntimeError(f"Word COM conversion failed: {e}")
        finally:
            if word:
                word.Quit()

    def _convert_with_docx2pdf(self, docx_path: Path, pdf_path: Path) -> Path:
        """Convert using docx2pdf library."""
        try:
            from docx2pdf import convert
            import signal

            def timeout_handler(signum, frame):
                raise TimeoutError("PDF conversion timed out")

            pdf_path.parent.mkdir(parents=True, exist_ok=True)

            # Set timeout for conversion (60 seconds for permission dialogs)
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(60)

            try:
                print(f"Converting to PDF using Word... (may request permissions)")
                convert(str(docx_path), str(pdf_path))
                signal.alarm(0)  # Cancel the alarm
                return pdf_path
            except TimeoutError:
                signal.alarm(0)
                raise RuntimeError("PDF conversion timed out after 60 seconds. This may be due to permission dialogs.")

        except Exception as e:
            signal.alarm(0)  # Ensure alarm is cancelled
            if "permission" in str(e).lower() or "access" in str(e).lower():
                raise RuntimeError(f"PDF conversion failed due to permissions. Please grant file access to Word in System Preferences > Security & Privacy > Files and Folders. Error: {e}")
            else:
                raise RuntimeError(f"docx2pdf conversion failed: {e}")

    def _convert_with_pandoc(self, docx_path: Path, pdf_path: Path) -> Path:
        """Convert using Pandoc."""
        try:
            pdf_path.parent.mkdir(parents=True, exist_ok=True)

            # Try different PDF engines in order of preference
            pdf_engines = ["xelatex", "pdflatex", "wkhtmltopdf"]

            for engine in pdf_engines:
                try:
                    result = subprocess.run([
                        "pandoc",
                        str(docx_path),
                        "-o", str(pdf_path),
                        f"--pdf-engine={engine}",
                        "--quiet"
                    ],
                    capture_output=True,
                    text=True,
                    timeout=120
                    )

                    if result.returncode == 0:
                        return pdf_path
                    else:
                        # If this engine failed, try the next one
                        continue

                except subprocess.TimeoutExpired:
                    continue

            # If all engines failed, try without specifying engine (use default)
            result = subprocess.run([
                "pandoc",
                str(docx_path),
                "-o", str(pdf_path),
                "--quiet"
            ],
            capture_output=True,
            text=True,
            timeout=120
            )

            if result.returncode != 0:
                raise RuntimeError(f"Pandoc conversion failed: No suitable PDF engine found. Install LaTeX (xelatex/pdflatex) or wkhtmltopdf for better Pandoc PDF support. Error: {result.stderr}")

            return pdf_path

        except subprocess.TimeoutExpired:
            raise RuntimeError("Pandoc conversion timed out")
        except Exception as e:
            raise RuntimeError(f"Pandoc conversion failed: {e}. Consider installing LaTeX or wkhtmltopdf for PDF support.")

    def get_available_method(self) -> str:
        """Get the name of available conversion method."""
        return self.conversion_method