def test_openpyxl_pagefield_import():
    """Verify that PageField can be imported from openpyxl."""
    from openpyxl.pivot.table import PageField
    assert PageField is not None
