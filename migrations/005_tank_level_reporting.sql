-- ====================================================================
-- Migration: 005_tank_level_reporting.sql
-- Description: Isolated tables, constraints, seeded factors and 13K17VEN
--              chart for Daily Tank Level Reporting.
-- Note: Completely additive. Does NOT alter any existing tables.
-- ====================================================================

-- 1. Tank Models Table
CREATE TABLE IF NOT EXISTS tank_models (
    id SERIAL PRIMARY KEY,
    model_name VARCHAR(100) UNIQUE NOT NULL,
    description TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Tank Level Gauge Chart Calibration Rows
CREATE TABLE IF NOT EXISTS tank_chart_rows (
    id SERIAL PRIMARY KEY,
    model_id INT NOT NULL REFERENCES tank_models(id) ON DELETE CASCADE,
    gas VARCHAR(50) NOT NULL,
    mmwc NUMERIC(10,2) NOT NULL,
    litres NUMERIC(12,2) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_tank_chart_row UNIQUE (model_id, gas, mmwc)
);
CREATE INDEX IF NOT EXISTS idx_tank_chart_rows_lookup ON tank_chart_rows(model_id, gas, mmwc);

-- 3. Gas Conversion Factors Table
CREATE TABLE IF NOT EXISTS tank_gas_factors (
    id SERIAL PRIMARY KEY,
    gas VARCHAR(50) UNIQUE NOT NULL,
    kg_per_litre NUMERIC(10,4) NOT NULL,
    sm3_per_kg NUMERIC(10,4) NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 4. Storage Tanks Table (Admin Setup)
CREATE TABLE IF NOT EXISTS storage_tanks (
    id SERIAL PRIMARY KEY,
    name VARCHAR(100) UNIQUE NOT NULL,
    gas VARCHAR(50) NOT NULL,
    model_id INT NOT NULL REFERENCES tank_models(id),
    location_notes TEXT,
    display_order INT DEFAULT 0,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_storage_tanks_active ON storage_tanks(is_active, display_order);

-- 5. Daily Tank Readings Table
CREATE TABLE IF NOT EXISTS tank_readings (
    id SERIAL PRIMARY KEY,
    tank_id INT NOT NULL REFERENCES storage_tanks(id),
    reading_date DATE NOT NULL,
    reading_time TIME,
    mmwc NUMERIC(10,2) NOT NULL,
    pressure NUMERIC(10,2),
    pressure_unit VARCHAR(20) DEFAULT 'kg/cm²',
    pressure_bar NUMERIC(10,4),
    litres NUMERIC(14,4),
    kg NUMERIC(14,4),
    sm3 NUMERIC(14,4),
    chart_model_used VARCHAR(100) NOT NULL,
    kg_per_litre_used NUMERIC(10,4) NOT NULL,
    sm3_per_kg_used NUMERIC(10,4) NOT NULL,
    net_change_kg NUMERIC(14,4),
    status VARCHAR(50) DEFAULT 'ok', -- 'ok', 'below_chart', 'above_full'
    recorded_by_id INT REFERENCES users(id),
    notes TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_tank_daily_reading UNIQUE (tank_id, reading_date)
);
CREATE INDEX IF NOT EXISTS idx_tank_readings_date ON tank_readings(reading_date, tank_id);
CREATE INDEX IF NOT EXISTS idx_tank_readings_tank_date ON tank_readings(tank_id, reading_date DESC);

-- ====================================================================
-- SEED DATA
-- ====================================================================

-- Seed Tank Model 13K17VEN
INSERT INTO tank_models (model_name, description)
VALUES ('13K17VEN', 'Cryogenic storage tank 11,623L capacity (Full levels: LIN 4461, LOX 6299, LAR 7709 mmWC)')
ON CONFLICT (model_name) DO NOTHING;

-- Seed Gas Conversion Factors
INSERT INTO tank_gas_factors (gas, kg_per_litre, sm3_per_kg)
VALUES 
    ('Oxygen', 1.1420, 0.7700),
    ('Nitrogen', 0.8080, 0.8800),
    ('Argon', 1.3930, 0.6170)
ON CONFLICT (gas) DO NOTHING;

-- Seed 13K17VEN Chart Rows
DO $$
DECLARE
    m_id INT;
BEGIN
    SELECT id INTO m_id FROM tank_models WHERE model_name = '13K17VEN' LIMIT 1;
    IF m_id IS NOT NULL THEN
        -- Nitrogen rows
        INSERT INTO tank_chart_rows (model_id, gas, mmwc, litres) VALUES
            (m_id, 'Nitrogen', 400, 767),
            (m_id, 'Nitrogen', 800, 1838),
            (m_id, 'Nitrogen', 1200, 2910),
            (m_id, 'Nitrogen', 1600, 3981),
            (m_id, 'Nitrogen', 2000, 5053),
            (m_id, 'Nitrogen', 2400, 6124),
            (m_id, 'Nitrogen', 2800, 7196),
            (m_id, 'Nitrogen', 3200, 8267),
            (m_id, 'Nitrogen', 3600, 9338),
            (m_id, 'Nitrogen', 4000, 10410),
            (m_id, 'Nitrogen', 4461, 11623)
        ON CONFLICT (model_id, gas, mmwc) DO NOTHING;

        -- Oxygen rows
        INSERT INTO tank_chart_rows (model_id, gas, mmwc, litres) VALUES
            (m_id, 'Oxygen', 400, 460),
            (m_id, 'Oxygen', 800, 1213),
            (m_id, 'Oxygen', 1200, 1972),
            (m_id, 'Oxygen', 1600, 2730),
            (m_id, 'Oxygen', 2000, 3489),
            (m_id, 'Oxygen', 2400, 4248),
            (m_id, 'Oxygen', 2800, 5007),
            (m_id, 'Oxygen', 3200, 5765),
            (m_id, 'Oxygen', 3600, 6524),
            (m_id, 'Oxygen', 4000, 7283),
            (m_id, 'Oxygen', 4400, 8042),
            (m_id, 'Oxygen', 4800, 8800),
            (m_id, 'Oxygen', 5200, 9559),
            (m_id, 'Oxygen', 5600, 10318),
            (m_id, 'Oxygen', 6000, 11076),
            (m_id, 'Oxygen', 6299, 11623)
        ON CONFLICT (model_id, gas, mmwc) DO NOTHING;

        -- Argon rows
        INSERT INTO tank_chart_rows (model_id, gas, mmwc, litres) VALUES
            (m_id, 'Argon', 400, 330),
            (m_id, 'Argon', 800, 936),
            (m_id, 'Argon', 1200, 1556),
            (m_id, 'Argon', 1600, 2176),
            (m_id, 'Argon', 2000, 2796),
            (m_id, 'Argon', 2400, 3416),
            (m_id, 'Argon', 2800, 4036),
            (m_id, 'Argon', 3200, 4656),
            (m_id, 'Argon', 3600, 5276),
            (m_id, 'Argon', 4000, 5896),
            (m_id, 'Argon', 4400, 6516),
            (m_id, 'Argon', 4800, 7136),
            (m_id, 'Argon', 5200, 7756),
            (m_id, 'Argon', 5600, 8376),
            (m_id, 'Argon', 6000, 8996),
            (m_id, 'Argon', 6400, 9616),
            (m_id, 'Argon', 6800, 10236),
            (m_id, 'Argon', 7200, 10856),
            (m_id, 'Argon', 7600, 11476),
            (m_id, 'Argon', 7709, 11623)
        ON CONFLICT (model_id, gas, mmwc) DO NOTHING;
    END IF;
END $$;
