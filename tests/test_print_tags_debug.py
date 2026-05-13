from openpyxl.pivot.table import TableDefinition, RowColField, Location

def test_pivot_xml_tags():
    """Verify the XML tag names for pivot table definitions."""
    loc = Location(ref="A1:B2", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
    td = TableDefinition(name="T", cacheId=0, dataCaption="C", location=loc, colFields=[RowColField(x=-2)])
    xml = td.to_tree()
    assert xml.tag == "pivotTableDefinition"
    child_tags = [c.tag for c in xml]
    assert "location" in child_tags
    assert "colFields" in child_tags
