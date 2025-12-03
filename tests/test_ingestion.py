import dlt

from ea_dagster.assets.ingestion import get_pipeline


class DummyPipeline:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs


def test_get_pipeline_sets_must_attach_to_local_pipeline(monkeypatch):
    called = {}

    def fake_pipeline(*args, **kwargs):
        called["args"] = args
        called["kwargs"] = kwargs
        return DummyPipeline(*args, **kwargs)

    monkeypatch.setattr(dlt, "pipeline", fake_pipeline)

    p = get_pipeline()
    assert isinstance(p, DummyPipeline)
    assert "pipeline_name" in called["kwargs"]
    assert "ea_ingest__staging" in called["kwargs"]["pipeline_name"]
    # ensure the important flag is passed as False
    assert "must_attach_to_local_pipeline" in called["kwargs"]
    assert called["kwargs"]["must_attach_to_local_pipeline"] is False

    # Second call should return cached pipeline and not re-call dlt.pipeline
    called.clear()
    p2 = get_pipeline()
    assert p is p2
    assert called == {}
