from docx import Document
from docx.shared import Mm, Pt, Inches


document = Document()


# Change Section sizes
# Mm, Pt, Inches
first_section = document.sections[0]
first_section.left_margin = Mm(0)
first_section.right_margin = Mm(0)
first_section.top_margin = Mm(0)
first_section.bottom_margin = Mm(0)
first_section.header_distance = Mm(0)
first_section.footer_distance = Mm(0)


# Access section Header
header = document.sections[0].header
