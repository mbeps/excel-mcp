from openpyxl.pivot.table import TableDefinition, Location

def test_openpyxl_empty_list_pivot_fields():
    """Verify that openpyxl TableDefinition accepts empty lists for rowFields."""
    loc = Location(ref="A1:B2", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
    td = TableDefinition(name="T", cacheId=0, dataCaption="C", location=loc, rowFields=[])
    assert td.rowFields == []
