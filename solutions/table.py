# Add table after specific paragraph
# Original Answer https://github.com/python-openxml/python-docx/issues/156#issuecomment-77674193
def move_table_after(table, paragraph):
    tbl, p = table._tbl, paragraph._p
    p.addnext(tbl)
