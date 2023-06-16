from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# Add table after specific paragraph
# Original Answer https://github.com/python-openxml/python-docx/issues/156#issuecomment-77674193
def move_table_after(table, paragraph):
    tbl, p = table._tbl, paragraph._p
    p.addnext(tbl)


# Change cell borders in table
# Original Answer https://stackoverflow.com/a/49615968/10462999
# Specification for Table Cells http://officeopenxml.com/WPtableBorders.php
def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))


# Original Answer https://stackoverflow.com/a/55177526/10462999
def set_cell_margins(cell, **kwargs):
    """
    cell:  actual cell instance you want to modify

    usage:

        set_cell_margins(cell, top=50, start=50, bottom=50, end=50)

    provided values are in twentieths of a point (1/1440 of an inch).
    read more here: http://officeopenxml.com/WPtableCellMargins.php
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')

    for m in [
        "top",
        "start",
        "bottom",
        "end",
    ]:
        if m in kwargs:
            node = OxmlElement("w:{}".format(m))
            node.set(qn('w:w'), str(kwargs.get(m)))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)

    tcPr.append(tcMar)


# Set width for column
# Original Answer https://stackoverflow.com/a/43053996/10462999
def set_col_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    table_cells = table._cells
    for row_index in range(len(table.rows)):
        row_cells = table_cells[row_index*2:(row_index+1)*2]
        row_cells[0].width = widths[0]
        row_cells[1].width = widths[1]

    for col, width in zip(table.columns, widths):  # some editors respect per cell width, another per col width
        col.width = width


def insert_table_properties(table, prop_element, **kwargs):
    """
    insert_table_properties(
        table,
        "tblPr",
        tblStyle={"val": "a"},
        tblW={"w": "8640", "type": "dxa"}
    )
    """

    tbl_pr = table._element.xpath(f"w:{prop_element}")
    if not tbl_pr:
        return
    for key, values in kwargs.items():
        elem = OxmlElement(f"w:{key}")
        for prop, value in values.items():
            elem.set(qn(f"w:{prop}"), str(value))

        tbl_pr[0].append(elem)


# Original Answer: https://github.com/python-openxml/python-docx/issues/322#issuecomment-265018856
def set_repeat_table_header(row):
    """ set repeat table row on every new page
    """
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)
    return row
