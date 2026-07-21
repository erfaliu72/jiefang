import os

# Load .env file for environment variables
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if os.path.exists(_env_path):
    with open(_env_path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                k, _, v = line.partition('=')
                os.environ.setdefault(k.strip(), v.strip())

from app import app, get_db, migrate_legacy_pending_approvals, run_daily_collect, start_scheduler
from database import init_db, seed_data


def bootstrap_runtime():
    try:
        init_db()
    except Exception as exc:
        print(f'[bootstrap] init_db error: {exc}')
    try:
        migration_conn = get_db()
        try:
            migrate_legacy_pending_approvals(migration_conn)
            migration_conn.commit()
        finally:
            migration_conn.close()
    except Exception as exc:
        print(f'[bootstrap] migration error: {exc}')
    try:
        seed_data()
    except Exception as exc:
        print(f'[bootstrap] seed_data error: {exc}')
    try:
        run_daily_collect(force=False)
    except Exception as exc:
        print(f'[bootstrap] daily-collect startup error: {exc}')
    try:
        start_scheduler()
    except Exception as exc:
        print(f'[bootstrap] scheduler error: {exc}')
    print('[bootstrap] initialization complete')


bootstrap_runtime()

application = app
