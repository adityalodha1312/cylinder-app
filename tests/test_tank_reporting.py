import pytest
from services.tank_service import (
    calculate_tank_values,
    interpolate_litres,
    get_chart_rows,
    normalize_pressure,
    check_pressure_warning,
    compute_net_change
)


def test_oxygen_2200_mmwc():
    """Oxygen 2200 mmWC: 3868.5 L, about 4418 kg, about 3402 SM³."""
    res = calculate_tank_values('Oxygen', 2200, pressure=8.5, pressure_unit='kg/cm²')
    assert res['status'] == 'ok'
    assert pytest.approx(res['litres'], 0.01) == 3868.5
    assert pytest.approx(res['kg'], 0.1) == 4417.83
    assert res['kg_display'] == 4418
    assert pytest.approx(res['sm3'], 0.1) == 3401.73
    assert res['sm3_display'] == 3402


def test_oxygen_400_mmwc():
    """Oxygen 400 mmWC: 460 L, about 525 kg, about 405 SM³."""
    res = calculate_tank_values('Oxygen', 400, pressure=10.0, pressure_unit='kg/cm²')
    assert res['status'] == 'ok'
    assert res['litres'] == 460.0
    assert pytest.approx(res['kg'], 0.1) == 525.32
    assert res['kg_display'] == 525
    assert pytest.approx(res['sm3'], 0.1) == 404.50
    assert res['sm3_display'] in (404, 405)


def test_nitrogen_4461_mmwc():
    """Nitrogen 4461 mmWC: 11,623 L, about 9391 kg, about 8264 SM³."""
    res = calculate_tank_values('Nitrogen', 4461, pressure=12.0, pressure_unit='kg/cm²')
    assert res['status'] == 'ok'
    assert res['litres'] == 11623.0
    assert pytest.approx(res['kg'], 0.1) == 9391.38
    assert res['kg_display'] == 9391
    assert pytest.approx(res['sm3'], 0.1) == 8264.42
    assert res['sm3_display'] == 8264


def test_nitrogen_5000_above_full():
    """Nitrogen 5000 mmWC: full capacity (11,623 L) with an above_full flag."""
    res = calculate_tank_values('Nitrogen', 5000, pressure=15.0, pressure_unit='kg/cm²')
    assert res['status'] == 'above_full'
    assert res['litres'] == 11623.0
    assert res['kg_display'] == 9391
    assert res['sm3_display'] == 8264
    assert any("Above full level" in w for w in res['warnings'])


def test_oxygen_350_below_chart():
    """Oxygen 350 mmWC: below_chart, no kg or SM³."""
    res = calculate_tank_values('Oxygen', 350, pressure=5.0, pressure_unit='kg/cm²')
    assert res['status'] == 'below_chart'
    assert res['litres'] is None
    assert res['kg'] is None
    assert res['sm3'] is None
    assert res['kg_display'] is None
    assert res['sm3_display'] is None
    assert any("Below chart range" in w for w in res['warnings'])


def test_argon_exact_and_interpolation():
    """Argon 400, 2000, and 7709 mmWC."""
    # 400 mmWC
    res400 = calculate_tank_values('Argon', 400)
    assert res400['status'] == 'ok'
    assert res400['litres'] == 330.0

    # 2200 mmWC (between 2000=2796 and 2400=3416 -> 2796 + 0.5*620 = 3106 L)
    res2200 = calculate_tank_values('Argon', 2200)
    assert res2200['status'] == 'ok'
    assert pytest.approx(res2200['litres'], 0.01) == 3106.0
    assert pytest.approx(res2200['kg'], 0.1) == 3106.0 * 1.393
    assert pytest.approx(res2200['sm3'], 0.1) == (3106.0 * 1.393) * 0.617

    # 7709 mmWC (full level)
    res_full = calculate_tank_values('Argon', 7709)
    assert res_full['status'] == 'ok'
    assert res_full['litres'] == 11623.0


def test_invalid_and_negative_mmwc():
    """Reject non-numeric or negative mmWC, warn above 10,000."""
    res_neg = calculate_tank_values('Oxygen', -50)
    assert res_neg['status'] == 'error'

    res_nan = calculate_tank_values('Oxygen', 'abc')
    assert res_nan['status'] == 'error'

    res_high = calculate_tank_values('Oxygen', 12000)
    assert any("> 10,000" in w for w in res_high['warnings'])


def test_pressure_normalization():
    """Normalize pressure to bar and check warning thresholds."""
    # 10 kg/cm²
    p_bar1 = normalize_pressure(10, 'kg/cm²')
    assert pytest.approx(p_bar1, 0.01) == 9.81

    # 100 psi
    p_bar2 = normalize_pressure(100, 'psi')
    assert pytest.approx(p_bar2, 0.01) == 6.89

    # Warnings
    assert check_pressure_warning(45, 'kg/cm²') is not None
    assert check_pressure_warning(30, 'kg/cm²') is None
    assert check_pressure_warning(650, 'psi') is not None
    assert check_pressure_warning(500, 'psi') is None


def test_net_change_labels():
    """Verify net change formatting."""
    class MockReading:
        def __init__(self, kg):
            self.kg = kg

    # First reading
    diff, label = compute_net_change(4418, None)
    assert diff is None
    assert label == "First reading"

    # Used
    diff, label = compute_net_change(4000, MockReading(4418))
    assert pytest.approx(diff, 0.1) == -418
    assert label == "Used 418 kg (net)"

    # Filled
    diff, label = compute_net_change(6500, MockReading(4000))
    assert pytest.approx(diff, 0.1) == 2500
    assert label == "Tank filled +2,500 kg (net)"

    # No change
    diff, label = compute_net_change(4000, MockReading(4000))
    assert diff == 0
    assert label == "No change (net)"


def test_isolated_metadata_not_in_db_metadata():
    """Verify new tables are completely isolated from db.metadata so db.create_all() ignores them."""
    from app import db
    from models import StorageTank, TankReading, TankModel, TankChartRow, TankGasFactor

    assert 'storage_tanks' not in db.metadata.tables
    assert 'tank_readings' not in db.metadata.tables
    assert 'tank_models' not in db.metadata.tables
    assert 'tank_chart_rows' not in db.metadata.tables
    assert 'tank_gas_factors' not in db.metadata.tables


def test_api_tank_calculate_endpoint():
    """Test /api/tanks/calculate via Flask test client."""
    from app import app
    client = app.test_client()
    with client.session_transaction() as sess:
        sess['user'] = {'username': 'testdriver', 'role': 'driver', 'id': 1}

    resp = client.post('/api/tanks/calculate', json={
        'gas': 'Oxygen',
        'mmwc': 2200,
        'pressure': 8.5,
        'pressure_unit': 'kg/cm²',
        'model_name': '13K17VEN'
    })
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['status'] == 'ok'
    assert data['kg_display'] == 4418
    assert data['sm3_display'] == 3402
