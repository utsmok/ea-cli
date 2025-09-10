import polars as pl

from easy_access.settings import Settings


def test_store_complete_data_missing_material_id_standalone(tmp_path):
    """Standalone test: store_complete_data should raise ValueError when material_id is missing."""
    settings = Settings()
    from easy_access.sheets.sheet import store_complete_data

    out_file = tmp_path / "out.xlsx"
    data = pl.DataFrame({"title": ["Item 1"]})

    try:
        store_complete_data(settings, out_file, data)
        raised = False
    except ValueError as e:
        raised = True
        assert "Export dataframe missing required columns" in str(e)

    assert raised
