from src.validators import validate_all


def test_all_data_checks_pass():
    report = validate_all()
    assert report["status"] == "PASS"
    assert len(report["checks"]) >= 15


def test_rounding_differences_are_documented():
    report = validate_all()
    check = next(item for item in report["checks"] if item["name"] == "主营业务+其他业务=营业收入")
    assert check["max_difference"] == 0.01

