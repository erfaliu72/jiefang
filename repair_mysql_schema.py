"""Apply the non-destructive MySQL schema repair required by the SKU rollout."""

import os
from pathlib import Path


def load_env_file():
    env_path = Path(__file__).with_name(".env")
    if not env_path.exists():
        return
    for raw_line in env_path.read_text().splitlines():
        line = raw_line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip())


load_env_file()

if not os.environ.get("JJY_DB_HOST"):
    raise RuntimeError("JJY_DB_HOST must be configured to repair the MySQL schema")

from database import drop_legacy_model_guidance_price_unique_indexes, get_db


def run_statement(conn, sql):
    try:
        conn.execute(sql)
    except Exception as exc:
        error_code = exc.args[0] if getattr(exc, "args", None) else None
        if error_code not in (1050, 1060, 1061):
            raise


def main():
    conn = get_db()
    try:
        statements = [
            "ALTER TABLE vehicles ADD COLUMN `condition` VARCHAR(191)",
            "ALTER TABLE vehicles ADD COLUMN brand TEXT",
            "ALTER TABLE vehicles ADD COLUMN battery_capacity TEXT",
            "ALTER TABLE vehicles ADD COLUMN horsepower TEXT",
            "ALTER TABLE vehicles ADD COLUMN box_type VARCHAR(191)",
            "ALTER TABLE vehicles ADD COLUMN gear_position TEXT",
            "ALTER TABLE vehicles ADD COLUMN tailgate VARCHAR(191)",
            "ALTER TABLE vehicles ADD COLUMN battery_brand TEXT",
            """
            ALTER TABLE data_dictionaries
            MODIFY COLUMN category VARCHAR(191) NOT NULL,
            MODIFY COLUMN value VARCHAR(191) NOT NULL,
            MODIFY COLUMN energy_type VARCHAR(191)
            """,
            """
            ALTER TABLE skus
            MODIFY COLUMN car_type VARCHAR(191) NOT NULL,
            MODIFY COLUMN `condition` VARCHAR(191) DEFAULT '新车',
            MODIFY COLUMN box_type VARCHAR(191),
            MODIFY COLUMN tailgate VARCHAR(191) DEFAULT '无'
            """,
            "ALTER TABLE model_guidance_prices ADD COLUMN is_new VARCHAR(191) DEFAULT '新车'",
            "ALTER TABLE model_guidance_prices ADD COLUMN lease_deposit_guidance DOUBLE DEFAULT 0",
            "ALTER TABLE model_guidance_prices ADD COLUMN box_standard_price DOUBLE DEFAULT 0",
            "ALTER TABLE model_guidance_prices ADD COLUMN box_wide_price DOUBLE DEFAULT 0",
            "ALTER TABLE model_guidance_prices ADD COLUMN box_high_rail_price DOUBLE DEFAULT 0",
            "ALTER TABLE model_guidance_prices ADD COLUMN box_refrigerated_price DOUBLE DEFAULT 0",
            "ALTER TABLE model_guidance_prices ADD COLUMN box_flatbed_price DOUBLE DEFAULT 0",
            "ALTER TABLE model_guidance_prices ADD COLUMN tail_plate_price DOUBLE DEFAULT 0",
            "ALTER TABLE finance_plans ADD COLUMN effective_date TEXT",
            "ALTER TABLE finance_plans ADD COLUMN expiry_date TEXT",
            """
            ALTER TABLE finance_plans
            MODIFY COLUMN car_type VARCHAR(191),
            MODIFY COLUMN `condition` VARCHAR(191) DEFAULT '新车',
            MODIFY COLUMN box_type VARCHAR(191) DEFAULT '',
            MODIFY COLUMN status VARCHAR(191)
            """,
            """
            CREATE TABLE IF NOT EXISTS contract_vehicles (
                id BIGINT AUTO_INCREMENT PRIMARY KEY,
                contract_id BIGINT NOT NULL,
                vehicle_id BIGINT NOT NULL,
                is_primary BIGINT DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            """,
            "CREATE INDEX idx_data_dictionaries_cat ON data_dictionaries(category, status)",
            "CREATE UNIQUE INDEX idx_skus_comb ON skus(car_type, `condition`, box_type, tailgate)",
            "CREATE INDEX idx_contract_vehicles_contract ON contract_vehicles(contract_id)",
            "CREATE UNIQUE INDEX idx_contract_vehicles_uniq ON contract_vehicles(contract_id, vehicle_id)",
            "CREATE INDEX idx_finance_plans_scope ON finance_plans(car_type, `condition`, box_type, status)",
        ]
        for statement in statements:
            run_statement(conn, statement)
        conn.execute("""
            UPDATE model_guidance_prices
            SET is_new='新车'
            WHERE is_new IS NULL OR TRIM(is_new)=''
        """)
        conn.execute("""
            DELETE FROM model_guidance_prices
            WHERE id NOT IN (
                SELECT keep_id
                FROM (
                    SELECT MAX(id) AS keep_id
                    FROM model_guidance_prices
                    GROUP BY car_type, COALESCE(is_new, '新车')
                ) AS guidance_price_keep
            )
        """)
        drop_legacy_model_guidance_price_unique_indexes(conn.cursor())
        run_statement(
            conn,
            """
            CREATE UNIQUE INDEX idx_model_guidance_prices_car_type_is_new
            ON model_guidance_prices(car_type, is_new)
            """,
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
