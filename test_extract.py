from datetime import date
import pytest, make_sample, extract
from pathlib import Path
S = Path(__file__).parent / "samples"

@pytest.fixture(scope="module", autouse=True)
def samples(): make_sample.main()

def test_fields(tmp_path):
    f, rows = extract.main([str(S/"purchase_agreement_sample.pdf"), "-o", str(tmp_path/"o.xlsx"), "--accepted", "2026-10-22"])
    assert f["Purchase Price"] == "$289,500.00" and f["Closing"] == "30 days" and all(f.values())
    assert (tmp_path/"o.xlsx").exists()

def test_weekend_roll():
    assert extract.roll_weekend(date(2026, 10, 24)) == date(2026, 10, 26)  # Sat
    assert extract.roll_weekend(date(2026, 10, 25)) == date(2026, 10, 26)  # Sun
    assert extract.roll_weekend(date(2026, 10, 23)) == date(2026, 10, 23)

def test_protected(tmp_path):
    with pytest.raises(SystemExit):
        extract.main([str(S/"purchase_agreement_protected.pdf"), "-o", str(tmp_path/"x.xlsx")])
    f, _ = extract.main([str(S/"purchase_agreement_protected.pdf"), "-o", str(tmp_path/"x.xlsx"), "--password", "demo1234"])
    assert f["Seller"] == "Pat Q. Sample"
