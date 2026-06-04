from app import app, get_db, migrate_legacy_pending_approvals, run_daily_collect, start_scheduler
from database import init_db, seed_data


def bootstrap_runtime():
    init_db()
    migration_conn = get_db()
    try:
        migrate_legacy_pending_approvals(migration_conn)
        migration_conn.commit()
    finally:
        migration_conn.close()
    seed_data()
    try:
        run_daily_collect(force=False)
    except Exception as exc:
        print(f'[daily-collect] startup catch-up error: {exc}')
    start_scheduler()


bootstrap_runtime()

application = app
