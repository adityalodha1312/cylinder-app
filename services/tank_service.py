import math
from datetime import datetime, date
from decimal import Decimal
import sqlalchemy as sa
from models import TankModel, TankChartRow, TankGasFactor, StorageTank, TankReading

# ─── Default Constants ────────────────────────────────────────────────────────
DEFAULT_FACTORS = {
    'Oxygen':   {'kg_per_litre': 1.1420, 'sm3_per_kg': 0.7700},
    'Nitrogen': {'kg_per_litre': 0.8080, 'sm3_per_kg': 0.8800},
    'Argon':    {'kg_per_litre': 1.3930, 'sm3_per_kg': 0.6170},
}

# 13K17VEN Plate Chart Seed Data (mmWC -> Litres)
CHART_13K17VEN = {
    'Nitrogen': [
        (400.0, 767.0),
        (800.0, 1838.0),
        (1200.0, 2910.0),
        (1600.0, 3981.0),
        (2000.0, 5053.0),
        (2400.0, 6124.0),
        (2800.0, 7196.0),
        (3200.0, 8267.0),
        (3600.0, 9338.0),
        (4000.0, 10410.0),
        (4461.0, 11623.0),
    ],
    'Oxygen': [
        (400.0, 460.0),
        (800.0, 1213.0),
        (1200.0, 1972.0),
        (1600.0, 2730.0),
        (2000.0, 3489.0),
        (2400.0, 4248.0),
        (2800.0, 5007.0),
        (3200.0, 5765.0),
        (3600.0, 6524.0),
        (4000.0, 7283.0),
        (4400.0, 8042.0),
        (4800.0, 8800.0),
        (5200.0, 9559.0),
        (5600.0, 10318.0),
        (6000.0, 11076.0),
        (6299.0, 11623.0),
    ],
    'Argon': [
        (400.0, 330.0),
        (800.0, 936.0),
        (1200.0, 1556.0),
        (1600.0, 2176.0),
        (2000.0, 2796.0),
        (2400.0, 3416.0),
        (2800.0, 4036.0),
        (3200.0, 4656.0),
        (3600.0, 5276.0),
        (4000.0, 5896.0),
        (4400.0, 6516.0),
        (4800.0, 7136.0),
        (5200.0, 7756.0),
        (5600.0, 8376.0),
        (6000.0, 8996.0),
        (6400.0, 9616.0),
        (6800.0, 10236.0),
        (7200.0, 10856.0),
        (7600.0, 11476.0),
        (7709.0, 11623.0),
    ]
}

PRESSURE_CONVERSIONS = {
    'kg/cm²': 0.980665,
    'psi':    0.0689476,
    'bar':    1.0
}


def normalize_pressure(value, unit):
    """Normalize pressure to bar."""
    if value is None:
        return None
    try:
        val = float(value)
    except (ValueError, TypeError):
        return None
    factor = PRESSURE_CONVERSIONS.get(unit or 'kg/cm²', 0.980665)
    return round(val * factor, 4)


def check_pressure_warning(value, unit):
    """Warn if pressure is above normal gauge scale (40 kg/cm² / 600 psi)."""
    if value is None:
        return None
    try:
        val = float(value)
    except (ValueError, TypeError):
        return None
    u = (unit or 'kg/cm²').lower()
    if 'kg' in u and val > 40.0:
        return f"High pressure warning: {val} {unit} exceeds 40 kg/cm² gauge scale"
    if 'psi' in u and val > 600.0:
        return f"High pressure warning: {val} {unit} exceeds 600 psi gauge scale"
    if 'bar' in u and val > 39.2:
        return f"High pressure warning: {val} {unit} exceeds ~39 bar gauge scale"
    return None


def get_gas_factors(gas_name):
    """Retrieve active gas conversion factors from DB or default fallback."""
    try:
        row = TankGasFactor.query.filter(sa.func.lower(TankGasFactor.gas) == gas_name.lower()).first()
        if row:
            return float(row.kg_per_litre), float(row.sm3_per_kg)
    except Exception:
        pass
    default = DEFAULT_FACTORS.get(gas_name, DEFAULT_FACTORS.get(gas_name.capitalize(), {'kg_per_litre': 1.0, 'sm3_per_kg': 1.0}))
    return default['kg_per_litre'], default['sm3_per_kg']


def get_chart_rows(model_name, gas_name):
    """Retrieve chart calibration rows (mmwc, litres) sorted by mmwc."""
    try:
        m = TankModel.query.filter_by(model_name=model_name).first()
        if m:
            db_rows = TankChartRow.query.filter_by(model_id=m.id, gas=gas_name).order_by(TankChartRow.mmwc.asc()).all()
            if db_rows and len(db_rows) >= 2:
                return [(float(r.mmwc), float(r.litres)) for r in db_rows]
    except Exception:
        pass

    # Fallback to internal 13K17VEN chart
    canonical_gas = gas_name.capitalize()
    if canonical_gas in CHART_13K17VEN:
        return list(CHART_13K17VEN[canonical_gas])
    return []


def interpolate_litres(chart_rows, mmwc):
    """
    Look up reading in level gauge chart using proportional linear interpolation.
    Returns: (litres, status, message)
      status: 'ok', 'below_chart', 'above_full', 'error'
    """
    if not chart_rows or len(chart_rows) < 2:
        return None, 'error', "Chart data missing or incomplete"

    try:
        reading = float(mmwc)
    except (ValueError, TypeError):
        return None, 'error', "Invalid mmWC reading: must be numeric"

    if reading < 0:
        return None, 'error', "Invalid mmWC reading: negative numbers rejected"

    min_row = chart_rows[0]
    max_row = chart_rows[-1]

    # Rule: Below 400 mmWC (the first row), the chart has no data.
    if reading < min_row[0]:
        return None, 'below_chart', "Below chart range, tank very low"

    # Rule: Above the gas's full level, treat tank as full capacity (11,623 L).
    if reading > max_row[0]:
        return max_row[1], 'above_full', "Above full level, check the reading and gas"

    # Rule: Exact match on min or max
    if reading == min_row[0]:
        return min_row[1], 'ok', None
    if reading == max_row[0]:
        return max_row[1], 'ok', None

    # Proportional linear interpolation between surrounding rows
    for i in range(len(chart_rows) - 1):
        x1, L1 = chart_rows[i]
        x2, L2 = chart_rows[i + 1]

        if reading == x1:
            return L1, 'ok', None
        if reading == x2:
            return L2, 'ok', None

        if x1 < reading < x2:
            litres = L1 + ((reading - x1) / (x2 - x1)) * (L2 - L1)
            return litres, 'ok', None

    return max_row[1], 'above_full', "Above full level, check the reading and gas"


def calculate_tank_values(gas, mmwc, pressure=None, pressure_unit='kg/cm²', model_name='13K17VEN'):
    """
    Calculates litres, kg, SM³, pressure normalization, warnings and status.
    """
    warnings = []
    try:
        mmwc_val = float(mmwc)
        if mmwc_val > 10000.0:
            warnings.append(f"Reading {mmwc_val:g} mmWC is unusually high (> 10,000 mmWC)")
    except (ValueError, TypeError):
        return {
            'status': 'error',
            'error': 'Reading must be a valid numeric value',
            'litres': None,
            'kg': None,
            'sm3': None,
            'kg_display': None,
            'sm3_display': None,
            'warnings': ['Invalid mmWC number']
        }

    chart_rows = get_chart_rows(model_name, gas)
    litres, status, flag_msg = interpolate_litres(chart_rows, mmwc_val)
    if flag_msg:
        warnings.append(flag_msg)

    kg_factor, sm3_factor = get_gas_factors(gas)

    kg = None
    sm3 = None
    kg_disp = None
    sm3_disp = None

    if litres is not None and status in ('ok', 'above_full'):
        kg = litres * kg_factor
        sm3 = kg * sm3_factor
        kg_disp = round(kg)
        sm3_disp = round(sm3)

    p_bar = normalize_pressure(pressure, pressure_unit)
    p_warn = check_pressure_warning(pressure, pressure_unit)
    if p_warn:
        warnings.append(p_warn)

    return {
        'status': status,
        'litres': litres,
        'kg': kg,
        'sm3': sm3,
        'kg_display': kg_disp,
        'sm3_display': sm3_disp,
        'chart_model_used': model_name,
        'kg_per_litre_used': kg_factor,
        'sm3_per_kg_used': sm3_factor,
        'pressure_bar': p_bar,
        'warnings': warnings,
        'flag_message': flag_msg
    }


def get_previous_good_reading(tank_id, before_date):
    """
    Fetches the most recent earlier reading with a valid kg value (status 'ok' or 'above_full').
    """
    try:
        prev = (TankReading.query
                .filter(TankReading.tank_id == tank_id)
                .filter(TankReading.reading_date < before_date)
                .filter(TankReading.kg.isnot(None))
                .filter(TankReading.status.in_(['ok', 'above_full']))
                .order_by(TankReading.reading_date.desc(), TankReading.id.desc())
                .first())
        return prev
    except Exception:
        return None


def compute_net_change(current_kg, previous_reading):
    """
    Calculates net change since previous good reading.
    Returns: (change_kg, label)
    """
    if current_kg is None or previous_reading is None or previous_reading.kg is None:
        return None, "First reading"

    prev_kg = float(previous_reading.kg)
    diff = current_kg - prev_kg
    rounded_diff = round(diff)

    if rounded_diff < 0:
        label = f"Used {abs(rounded_diff):,} kg (net)"
    elif rounded_diff > 0:
        label = f"Tank filled +{rounded_diff:,} kg (net)"
    else:
        label = "No change (net)"

    return diff, label
