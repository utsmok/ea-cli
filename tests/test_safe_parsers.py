import pytest
from easy_access.db.update import safe_int, safe_float, safe_date, safe_enum, safe_compare_greater
from easy_access.db.models import Status
from datetime import datetime, date


def test_safe_int():
    assert safe_int("123") == 123
    assert safe_int(45.0) == 45
    assert safe_int(None) is None
    assert safe_int("notanumber") is None


def test_safe_float():
    assert safe_float("12.34") == pytest.approx(12.34)
    assert safe_float(5) == pytest.approx(5.0)
    assert safe_float(None) is None
    assert safe_float("abc") is None


def test_safe_date():
    assert safe_date("2020-01-02") == date(2020, 1, 2)
    dt = datetime(2021, 5, 4, 12, 0, 0)
    assert safe_date(dt) == date(2021, 5, 4)
    assert safe_date(None) is None
    assert safe_date("invalid") is None


def test_safe_enum():
    assert safe_enum(Status, "Published") == Status.PUBLISHED
    assert safe_enum(Status, "PUBLISHED") is None or isinstance(safe_enum(Status, "PUBLISHED"), Status)


def test_safe_compare_greater():
    assert safe_compare_greater(5, 3)
    assert not safe_compare_greater(2, 4)
    assert safe_compare_greater("10", "2")
    assert safe_compare_greater(date(2021,1,2), date(2020,12,31))
    assert not safe_compare_greater(None, 1)
