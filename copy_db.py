# -*- coding: utf-8 -*-
"""로컬 SQLite(msd.db) → 배포 DB(PostgreSQL) 데이터 복사

    set DATABASE_URL=postgresql://...   (Railway Postgres의 공개 접속 URL)
    py copy_db.py                      # 대상 DB가 비어 있을 때만 복사
    py copy_db.py --only-master        # 작업·단위작업·직원명단만 (회차·응답·조사표 제외)
    py copy_db.py --force              # 대상 테이블을 비우고 덮어쓰기 (주의)

앱 모델(db.metadata) 순서대로 복사하고, PostgreSQL이면 id 시퀀스를 맞춘다.
"""
import os, sys
from sqlalchemy import create_engine, MetaData, select, text

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = 'sqlite:///' + os.path.join(BASE, 'msd.db').replace(os.sep, '/')
MASTER = {'job', 'employee'}

if not os.environ.get('DATABASE_URL'):
    sys.exit('DATABASE_URL(대상 DB 주소)을 먼저 설정하세요.')

sys.path.insert(0, BASE)
import app as A          # 대상 DB에 테이블 생성(create_all) + 자동 마이그레이션

force = '--force' in sys.argv
only_master = '--only-master' in sys.argv
src = create_engine(SRC)
smeta = MetaData()
smeta.reflect(src)
ctx = A.app.app_context()
ctx.push()
dst = A.db.engine
is_pg = dst.url.get_backend_name().startswith('postgresql')

with src.connect() as sc, dst.begin() as dc:
    tables = [t for t in A.db.metadata.sorted_tables if not only_master or t.name in MASTER]
    if force:
        for t in reversed(tables):
            dc.execute(t.delete())
    for t in tables:
        if t.name not in smeta.tables:
            print(f'  - {t.name}: 원본에 없음, 건너뜀'); continue
        n_dst = dc.execute(select(A.db.func.count()).select_from(t)).scalar()
        if n_dst and not force:
            print(f'  - {t.name}: 대상에 이미 {n_dst}건 → 건너뜀 (--force 로 덮어쓰기)'); continue
        cols = [c.name for c in t.columns if c.name in smeta.tables[t.name].columns]
        rows = [dict(r._mapping) for r in sc.execute(select(*[smeta.tables[t.name].c[c] for c in cols]))]
        if rows:
            dc.execute(t.insert(), rows)
        if is_pg and 'id' in t.columns:
            dc.execute(text(f"SELECT setval(pg_get_serial_sequence('\"{t.name}\"','id'), COALESCE((SELECT MAX(id) FROM \"{t.name}\"), 1))"))
        print(f'  ✓ {t.name}: {len(rows)}건')
print('완료')
