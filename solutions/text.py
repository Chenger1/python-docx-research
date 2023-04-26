from docx.text.paragraph import Paragraph
from docx.oxml.xmlchemy import OxmlElement


# Original Answer: https://stackoverflow.com/questions/48663788/python-docx-insert-a-paragraph-after
def insert_paragraph_after(paragraph, text=None, style=None):
    """Insert a new paragraph after the given paragraph."""
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    if style is not None:
        new_para.style = style
    return new_para


# Insert paragraph at specific position
# You have to insert content in reverse order
def insert_at_position(text, position, base_paragraph):
    paragraph = insert_paragraph_after(base_paragraph, text)
    for i in range(position+1):
        insert_paragraph_after(base_paragraph, "")
    return paragraph


# Delete parapgraph
# Original answer https://github.com/python-openxml/python-docx/issues/33#issuecomment-77661907
# Read comment because solution has some caveats
def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
