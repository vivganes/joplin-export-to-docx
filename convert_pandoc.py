import os
import tempfile
import pypandoc

def convert_markdowns_to_docx(md_files=None, output_file="combined.docx", chapters_file=None, folder_name=None):
    """
    Convert markdown files to DOCX using Pandoc, handling <partname> logic from chapters.txt.
    Args:
        md_files: List of markdown files (used if chapters_file is None)
        output_file: Output DOCX file path
        chapters_file: Path to chapters.txt file (takes precedence over md_files)
        folder_name: Optional folder where markdown files are located
    """
    temp_files = []
    combined_md = ""
    if chapters_file and os.path.exists(chapters_file):
        with open(chapters_file, encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]
        for line in lines:
            if line.startswith("<partname>"):
                part_title = line.replace("<partname>", "").strip()
                combined_md += f"\n\n# {part_title}\n\n<!-- PAGEBREAK -->\n\n"
            else:
                chapter_file = line
                if folder_name and not os.path.isabs(chapter_file):
                    chapter_file = os.path.join(folder_name, chapter_file)
                if not chapter_file.endswith(".md"):
                    chapter_file += ".md"
                if os.path.exists(chapter_file):
                    with open(chapter_file, encoding="utf-8") as chf:
                        combined_md += chf.read() + "\n\n<!-- PAGEBREAK -->\n\n"
    elif md_files:
        for f in md_files:
            chapter_file = os.path.join(folder_name, f) if folder_name and not os.path.isabs(f) else f
            if os.path.exists(chapter_file):
                with open(chapter_file, encoding="utf-8") as chf:
                    combined_md += chf.read() + "\n\n:::pagebreak:::\n\n"
    else:
        raise ValueError("Either chapters_file must exist or md_files must be provided")

    # Suppress alt text captions for images by removing alt text from markdown
    import re
    def suppress_image_alt(md):
        # Replace ![alt](src) with ![](src)
        return re.sub(r'!\[[^\]]*\]\(([^\)]+)\)', r'![](\1)', md)
    def reset_caps_in_code(md):
        return re.sub(r'(\n```[a-zA-Z0-9_+-]*)', r'\1', md)
    combined_md = suppress_image_alt(combined_md)
    combined_md = reset_caps_in_code(combined_md)

    # Write combined markdown to a temp file
    temp_md_file = tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode="w", encoding="utf-8")
    temp_md_file.write(combined_md)
    temp_md_file.close()
    temp_files.append(temp_md_file.name)

    # Use pypandoc to convert, set resource_path for images, enable raw_tex for page breaks
    # Choose your Pygments theme here (e.g., monokai, tango, zenburn, haddock, etc.)
    pygments_theme = "tango"  # Use a supported theme for DOCX
    pandoc_args = [
        "-f", "markdown+raw_tex",
        f"--syntax-highlighting={pygments_theme}",
        "--filter", "pandoc-pagebreak.py"
    ]
    if folder_name:
        pandoc_args.extend(["--resource-path", folder_name])
    try:
        pypandoc.convert_file(temp_md_file.name, 'docx', outputfile=output_file, extra_args=pandoc_args)
        print(f"Saved {output_file}")
    except Exception as e:
        print("Pandoc error:", e)

    # Clean up temp files
    for tf in temp_files:
        try:
            os.remove(tf)
        except Exception:
            pass

if __name__ == "__main__":
    output_docx = "book.docx"
    convert_markdowns_to_docx(output_file=output_docx, chapters_file="chapters.txt", folder_name="Test-First Copilot")
