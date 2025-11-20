import markdown
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Emu
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsmap
from pygments import highlight
from pygments.lexers import guess_lexer, get_lexer_by_name
from pygments.formatters.html import HtmlFormatter
from bs4 import BeautifulSoup
import re
import os
num_def_counter = [1]
def add_numbering_restart(document):
        numbering_part = document.part.numbering_part
        abstract_num_id = num_def_counter[0]
        num_id = num_def_counter[0]
        num_def_counter[0] += 1
        abstract_num_xml = f'''
        <w:abstractNum w:abstractNumId="{abstract_num_id}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:multiLevelType w:val="singleLevel"/>
            <w:lvl w:ilvl="0">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%1."/>
                <w:lvlJc w:val="left"/>
                <w:pPr/>
            </w:lvl>
        </w:abstractNum>
        '''
        abstract_num = parse_xml(abstract_num_xml)
        numbering_part.element.append(abstract_num)
        num_xml = f'<w:num w:numId="{num_id}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNumId w:val="{abstract_num_id}"/></w:num>'
        num = parse_xml(num_xml)
        numbering_part.element.append(num)
        return num_id
    # ...existing code...


def set_run_font(run, monospace=False, bold=False, italic=False, color=None, size=9):
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if monospace:
        run.font.name = "Courier New"
        rPr = run._element.rPr
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), "Courier New")
        rFonts.set(qn("w:hAnsi"), "Courier New")
        rPr.append(rFonts)
    if color:
        run.font.color.rgb = RGBColor(*color)


def parse_css_styles(formatter):
    """Return dict {class: {color, bold, italic}} from Pygments CSS."""
    css = formatter.get_style_defs('.highlight')
    styles = {}
    for rule in css.split("\n"):
        m = re.match(r"\.highlight\s+\.([a-zA-Z0-9_-]+)\s*\{([^}]*)\}", rule.strip())
        if not m:
            continue
        cls, body = m.groups()
        props = {}
        for part in body.split(";"):
            part = part.strip()
            if part.startswith("color:"):
                hexcol = part.split(":")[1].strip()
                if hexcol.startswith("#") and len(hexcol) == 7:
                    r = int(hexcol[1:3], 16)
                    g = int(hexcol[3:5], 16)
                    b = int(hexcol[5:7], 16)
                    props["color"] = (r, g, b)
            if "bold" in part:
                props["bold"] = True
            if "italic" in part:
                props["italic"] = True
        styles[cls] = props
    return styles


def set_cell_width(cell, emu_value):
    """Force table cell width in EMU."""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(emu_value)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


def add_code_block(document, code, theme="friendly", language=None):
    """Render syntax-highlighted code in a 2-column table (line numbers + code)."""
    # Choose lexer
    try:
        if language:
            lexer = get_lexer_by_name(language)
        else:
            lexer = guess_lexer(code)
    except Exception:
        lexer = get_lexer_by_name("text")

    # Format with chosen theme
    formatter = HtmlFormatter(nowrap=False, style=theme)
    highlighted = highlight(code, lexer, formatter)
    soup = BeautifulSoup(highlighted, "html.parser")
    css_styles = parse_css_styles(formatter)

    pre = soup.find("pre")
    if pre is None:
        return

    # Split into lines, keeping <span>
    lines, current_line = [], []
    for node in pre.children:
        if isinstance(node, str):
            parts = node.split("\n")
            for i, part in enumerate(parts):
                if i > 0:
                    lines.append(current_line)
                    current_line = []
                if part:
                    current_line.append(part)
        else:
            current_line.append(node)
    if current_line:
        lines.append(current_line)

    # Two-column table
    table = document.add_table(rows=0, cols=2)
    tbl_pr = table._element.tblPr
    tbl_layout = OxmlElement('w:tblLayout')
    tbl_layout.set(qn('w:type'), 'fixed')
    tbl_pr.append(tbl_layout)

    for i, parts in enumerate(lines, start=1):
        num_cell, code_cell = table.add_row().cells

        # Fix widths
        set_cell_width(num_cell, Emu(600))   # ~3 char width
        set_cell_width(code_cell, Emu(8000))

        # Line number
        para_num = num_cell.paragraphs[0]
        run_num = para_num.add_run(str(i))
        set_run_font(run_num, monospace=True, size=9, color=(128, 128, 128))
        para_num.alignment = 2
        para_num.paragraph_format.line_spacing = 1.2
        para_num.paragraph_format.space_after = Pt(0)

        # Code line
        para_code = code_cell.paragraphs[0]
        para_code.paragraph_format.line_spacing = 1.2
        para_code.paragraph_format.space_after = Pt(0)

        for node in parts:
            if isinstance(node, str):
                run = para_code.add_run(node)
                set_run_font(run, monospace=True, size=9)
            elif node.name == "span":
                text = node.get_text()
                run = para_code.add_run(text)
                set_run_font(run, monospace=True, size=9)
                for cls in node.get("class", []):
                    style = css_styles.get(cls, {})
                    if "color" in style:
                        run.font.color.rgb = RGBColor(*style["color"])
                    if style.get("bold"):
                        run.bold = True
                    if style.get("italic"):
                        run.italic = True

    # Grey background
    for row in table.rows:
        for cell in row.cells:
            tc_pr = cell._element.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), "F2F2F2")
            tc_pr.append(shd)

    # Add spacing after block
    spacer = document.add_paragraph()
    spacer.paragraph_format.space_before = Pt(6)
    spacer.paragraph_format.space_after = Pt(6)


def handle_inline(element, para):
    if element.name is None:
        para.add_run(str(element))
    elif element.name == "code":
        run = para.add_run(element.get_text())
        set_run_font(run, monospace=True, color=(200, 0, 0))
    elif element.name == "strong":
        run = para.add_run(element.get_text())
        run.bold = True
    elif element.name == "em":
        run = para.add_run(element.get_text())
        run.italic = True
    elif element.name == "a":
        run = para.add_run(element.get_text())
        run.font.color.rgb = RGBColor(0, 0, 255)
        run.underline = True
    else:
        for child in element.children:
            handle_inline(child, para)


def process_list(document, element, level=0, folder_name=None):
    num_id = None
    if element.name == "ol":
        document.add_paragraph()  # Insert blank paragraph to force restart
        num_id = add_numbering_restart(document)
    for idx, li in enumerate(element.find_all("li", recursive=False), start=1):
        style = "ListBullet" if element.name == "ul" else None
        if style:
            para = document.add_paragraph(style=style)
        else:
            para = document.add_paragraph()
            p = para._p
            numPr = OxmlElement('w:numPr')
            ilvl = OxmlElement('w:ilvl')
            ilvl.set(qn('w:val'), str(level))
            numPr.append(ilvl)
            numId = OxmlElement('w:numId')
            numId.set(qn('w:val'), str(num_id))
            numPr.append(numId)
            p.append(numPr)

        if level > 0:
            para.paragraph_format.left_indent = Cm(0.75 * level)

        for child in li.children:
            if isinstance(child, str) and child.strip() == "":
                continue
            if getattr(child, "name", None) in ["ul", "ol"]:
                process_list(document, child, level + 1, folder_name=folder_name)
            else:
                if getattr(child, "name", None) == "img":
                    run = para.add_run()
                    img_src = child["src"]
                    if folder_name and not os.path.isabs(img_src):
                        img_src = os.path.join(folder_name, img_src)
                    run.add_picture(img_src)
                    # Center align image
                    para.alignment = 1  # 1 = center
                else:
                    handle_inline(child, para)


def add_markdown_content(document, md_content, theme="friendly", folder_name=None):
    html = markdown.markdown(md_content, extensions=["fenced_code", "tables"])
    soup = BeautifulSoup(html, "html.parser")

    def process_element(element):
        if element.name in ["h1", "h2", "h3", "h4"]:
            level = int(element.name[1])
            document.add_heading(element.get_text(), level=level)

        elif element.name == "p":
            para = document.add_paragraph()
            for child in element.children:
                if child.name == "img":
                    run  = para.add_run()
                    img_src = child["src"]
                    if folder_name and not os.path.isabs(img_src):
                        img_src = os.path.join(folder_name, img_src)
                    run.add_picture(img_src, width=Cm(10))
                else:
                    handle_inline(child, para)

        elif element.name in ["ul", "ol"]:
            process_list(document, element, level=0, folder_name=folder_name)

        elif element.name == "table":
            rows = element.find_all("tr")
            n_cols = len(rows[0].find_all(["td", "th"]))
            table = document.add_table(rows=len(rows), cols=n_cols)
            for r, row in enumerate(rows):
                for c, cell in enumerate(row.find_all(["td", "th"])):
                    table.cell(r, c).text = cell.get_text(strip=True)
        
        elif element.name == "hr":
            para = document.add_paragraph()
            run = para.add_run()
            border = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '1')
            bottom.set(qn('w:space'), '0')
            bottom.set(qn('w:color'), 'auto')
            border.append(bottom)
            para._element.get_or_add_pPr().append(border)

        elif element.name == "pre":
            code_tag = element.find("code")
            if code_tag:
                lang = None
                # Look at classes, normalize them
                for cls in code_tag.get("class", []):
                    if cls.startswith("language-"):
                        lang = cls.replace("language-", "").lower()
                    elif cls.startswith("lang-"):
                        lang = cls.replace("lang-", "").lower()

                # Special-case common aliases
                if lang in ("py", "python3", "python"):
                    lang = "python"

                add_code_block(document, code_tag.get_text(), theme=theme, language=lang)

        else:
            if hasattr(element, "children"):
                for child in element.children:
                    if getattr(child, "name", None):
                        process_element(child)

    for element in soup.children:
        if getattr(element, "name", None):
            process_element(element)


def read_chapters_file(chapters_file, folder_name=None):
    """
    Read chapters.txt file and return list of items to process.
    Each item is either a markdown file path or a part name.
    """
    items = []
    with open(chapters_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            
            if line.startswith("<partname>"):
                # Extract part name after the tag
                part_name = line[10:].strip()  # Remove "<partname>" prefix
                items.append({"type": "part", "name": part_name})
            else:
                # It's a markdown file
                md_file = line
                if folder_name and not os.path.isabs(md_file):
                    md_file = os.path.join(folder_name, md_file)
                items.append({"type": "markdown", "path": md_file})
    
    return items


def add_part_page(document, part_name):
    """
    Add a page with just the part name text and a page break.
    """
    para = document.add_paragraph()
    run = para.add_run(part_name)
    run.font.size = Pt(16)
    run.bold = True
    para.alignment = 1  # Center alignment
    document.add_page_break()


def convert_markdowns_to_docx(md_files=None, output_file="combined.docx", theme="friendly", 
                              chapters_file="chapters.txt", folder_name=None):
    """
    Convert markdown files to DOCX.
    
    Args:
        md_files: List of markdown files (used if chapters_file is None)
        output_file: Output DOCX file path
        theme: Syntax highlighting theme
        chapters_file: Path to chapters.txt file (takes precedence over md_files)
        folder_name: Optional folder where markdown files are located
    """
    document = Document()
    
    # If chapters_file exists, use it; otherwise fall back to md_files
    if chapters_file and os.path.exists(chapters_file):
        items = read_chapters_file(chapters_file, folder_name)
        for i, item in enumerate(items):
            if item["type"] == "part":
                add_part_page(document, item["name"])
            elif item["type"] == "markdown":
                if i > 0:
                    document.add_section(1)  # WD_SECTION.NEW_PAGE
                if not item["path"].endswith(".md"):
                    item["path"] += ".md"
                with open(item["path"], "r", encoding="utf-8") as f:
                    md_content = f.read()
                add_markdown_content(document, md_content, theme=theme, folder_name=folder_name)
            print(f"Processed: {item}")
    elif md_files:
        for i, md_file in enumerate(md_files):
            if i > 0:
                document.add_section(1)  # WD_SECTION.NEW_PAGE
            full_path = md_file
            if folder_name and not os.path.isabs(md_file):
                full_path = os.path.join(folder_name, md_file)
            with open(full_path, "r", encoding="utf-8") as f:
                md_content = f.read()
            add_markdown_content(document, md_content, theme=theme, folder_name=folder_name)
    else:
        raise ValueError("Either chapters_file must exist or md_files must be provided")
    document.save(output_file)


if __name__ == "__main__":
    output_docx = "book.docx"
    convert_markdowns_to_docx(output_file=output_docx, theme="friendly", chapters_file="chapters.txt", folder_name="Test-First Copilot")
    print(f"Saved {output_docx}")