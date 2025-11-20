import os
import tempfile
import pytest
import pypandoc
from convert import convert_markdowns_to_docx

def test_raises_if_no_input(tmp_path):
    with pytest.raises(ValueError):
        convert_markdowns_to_docx()

def test_combines_md_files(tmp_path, monkeypatch):
    # Create two markdown files
    md1 = tmp_path / "a.md"
    md2 = tmp_path / "b.md"
    md1.write_text("# Chapter 1\nHello")
    md2.write_text("# Chapter 2\nWorld")
    output_docx = tmp_path / "out.docx"
    # Patch pypandoc.convert_file to avoid real conversion
    called = {}
    def fake_convert_file(input_file, to, outputfile=None, extra_args=None):
        called['input'] = input_file
        called['output'] = outputfile
        called['args'] = extra_args
        # Check that the temp file exists and contains both chapters
        with open(input_file, encoding="utf-8") as f:
            content = f.read()
            assert "Chapter 1" in content
            assert "Chapter 2" in content
            assert "Hello" in content
            assert "World" in content
        # Simulate output file creation
        with open(outputfile, "w") as outf:
            outf.write("fake docx")
        return None
    monkeypatch.setattr(pypandoc, "convert_file", fake_convert_file)
    convert_markdowns_to_docx(md_files=[str(md1), str(md2)], output_file=str(output_docx))
    assert output_docx.exists()
    assert called['output'] == str(output_docx)

def test_chapters_file_and_folder(tmp_path, monkeypatch):
    # Create folder and markdown files
    folder = tmp_path / "chapters"
    folder.mkdir()
    md1 = folder / "c1.md"
    md2 = folder / "c2.md"
    md1.write_text("# C1\nA")
    md2.write_text("# C2\nB")
    chapters_txt = tmp_path / "chapters.txt"
    chapters_txt.write_text("""
<partname> Part 1
c1.md
c2.md
""")
    output_docx = tmp_path / "out2.docx"
    called = {}
    def fake_convert_file(input_file, to, outputfile=None, extra_args=None):
        called['input'] = input_file
        called['output'] = outputfile
        with open(input_file, encoding="utf-8") as f:
            content = f.read()
            assert "Part 1" in content
            assert "# C1" in content
            assert "# C2" in content
        with open(outputfile, "w") as outf:
            outf.write("fake docx")
        return None
    monkeypatch.setattr(pypandoc, "convert_file", fake_convert_file)
    convert_markdowns_to_docx(output_file=str(output_docx), chapters_file=str(chapters_txt), folder_name=str(folder))
    assert output_docx.exists()
    assert called['output'] == str(output_docx)

def test_image_alt_suppression(tmp_path, monkeypatch):
    md = tmp_path / "img.md"
    md.write_text("![alt text](img.png)")
    output_docx = tmp_path / "img.docx"
    def fake_convert_file(input_file, to, outputfile=None, extra_args=None):
        with open(input_file, encoding="utf-8") as f:
            content = f.read()
            assert "![alt text]" not in content
            assert "![](img.png)" in content
        with open(outputfile, "w") as outf:
            outf.write("fake docx")
        return None
    monkeypatch.setattr(pypandoc, "convert_file", fake_convert_file)
    convert_markdowns_to_docx(md_files=[str(md)], output_file=str(output_docx))
    assert output_docx.exists()
