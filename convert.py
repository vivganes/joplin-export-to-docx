import markdown
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Emu
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pygments import highlight
from pygments.lexers import guess_lexer, get_lexer_by_name
from pygments.formatters.html import HtmlFormatter
from bs4 import BeautifulSoup
import re


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


def process_list(document, element, level=0):
    for idx, li in enumerate(element.find_all("li", recursive=False), start=1):
        style = "ListBullet" if element.name == "ul" else "ListNumber"
        para = document.add_paragraph(style=style)

        if level > 0:
            para.paragraph_format.left_indent = Cm(0.75 * level)

        for child in li.children:
            if isinstance(child, str) and child.strip() == "":
                continue
            if getattr(child, "name", None) in ["ul", "ol"]:
                process_list(document, child, level + 1)
            else:
                if child.name == "img":
                    print("Found image tag", child)
                    run = para.add_run()
                    run.add_picture(child["src"])
                else:
                    handle_inline(child, para)


def add_markdown_content(document, md_content, theme="friendly"):
    html = markdown.markdown(md_content, extensions=["fenced_code", "tables"])
    print("Converted Markdown to HTML \n ", html)
    soup = BeautifulSoup(html, "html.parser")

    def process_element(element):
        print("Processing element:", element.name)
        if element.name in ["h1", "h2", "h3", "h4"]:
            level = int(element.name[1])
            document.add_heading(element.get_text(), level=level)

        elif element.name == "p":
            para = document.add_paragraph()
            for child in element.children:
                if child.name == "img":
                    print("Found image tag", child)
                    run  = para.add_run()
                    run.add_picture(child["src"], width=Cm(10))
                else:
                    handle_inline(child, para)

        elif element.name in ["ul", "ol"]:
            process_list(document, element, level=0)

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


def convert_markdowns_to_docx(md_files, output_file, theme="friendly"):
    document = Document()
    for i, md_file in enumerate(md_files):
        with open(md_file, "r", encoding="utf-8") as f:
            md_content = f.read()
        add_markdown_content(document, md_content, theme=theme)
        if i < len(md_files) - 1:
            document.add_page_break()
    document.save(output_file)


if __name__ == "__main__":
    markdown_files = [ "chapter1.md"]
    output_docx = "combined.docx"
    convert_markdowns_to_docx(markdown_files, output_docx, theme="friendly")
    print(f"Saved {output_docx}")