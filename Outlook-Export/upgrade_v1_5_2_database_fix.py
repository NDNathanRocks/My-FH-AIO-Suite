from __future__ import annotations
from pathlib import Path
import py_compile
import shutil

HERE=Path(__file__).resolve().parent
CANDIDATES=[HERE/'inbox_email_exporter_v1_5_1.py',HERE/'inbox_email_exporter_v1_5.py']
SOURCE=next((p for p in CANDIDATES if p.exists()),None)
OUTPUT=HERE/'inbox_email_exporter_v1_5_2.py'

HISTORY=r'''def backup_and_migrate_database():
    """Upgrade any older exports table before a run can begin.

    Migration is additive, transactional, and backed up first. Older columns
    (`folder`, `size`) remain in place for compatibility and are copied into
    the newer audit columns (`source_folder`, `file_size`).
    """
    ensure_runtime()
    connection=sqlite3.connect(DB_PATH,timeout=30)
    try:
        connection.execute("PRAGMA busy_timeout=30000")
        integrity=connection.execute("PRAGMA integrity_check").fetchone()[0]
        if integrity!="ok":raise RuntimeError(f"Database integrity check failed before migration: {integrity}")
        exists=connection.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='exports'").fetchone()
        if not exists:
            connection.execute("""CREATE TABLE exports(
                store_id TEXT NOT NULL,entry_id TEXT NOT NULL,item_class INTEGER,
                item_type TEXT,source_folder TEXT,destination TEXT,file_size INTEGER,
                status TEXT,updated TEXT,error TEXT,PRIMARY KEY(store_id,entry_id))""")
            connection.execute("CREATE INDEX IF NOT EXISTS idx_exports_status ON exports(status)")
            connection.commit();return {"created":True,"added":[]}
        columns=[row[1] for row in connection.execute("PRAGMA table_info(exports)")]
        additions=[]
        required={"item_class":"INTEGER","item_type":"TEXT","source_folder":"TEXT","destination":"TEXT","file_size":"INTEGER","status":"TEXT","updated":"TEXT","error":"TEXT"}
        missing=[(name,kind) for name,kind in required.items() if name not in columns]
        if missing:
            stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
            backup=DB_PATH.with_name(f"outlook_export_pre_migration_{stamp}.backup.db")
            connection.commit();shutil.copy2(DB_PATH,backup)
            connection.execute("BEGIN IMMEDIATE")
            for name,kind in missing:
                connection.execute(f'ALTER TABLE exports ADD COLUMN "{name}" {kind}')
                additions.append(name)
            if "folder" in columns and "source_folder" in additions:
                connection.execute("UPDATE exports SET source_folder=folder WHERE source_folder IS NULL OR source_folder='' ")
            if "size" in columns and "file_size" in additions:
                connection.execute("UPDATE exports SET file_size=size WHERE file_size IS NULL")
            connection.execute("CREATE INDEX IF NOT EXISTS idx_exports_status ON exports(status)")
            integrity=connection.execute("PRAGMA integrity_check").fetchone()[0]
            if integrity!="ok":raise RuntimeError(f"Database integrity check failed during migration: {integrity}")
            connection.commit()
            return {"created":False,"added":additions,"backup":str(backup)}
        connection.execute("CREATE INDEX IF NOT EXISTS idx_exports_status ON exports(status)")
        connection.commit();return {"created":False,"added":[]}
    except Exception:
        connection.rollback();raise
    finally:connection.close()


class History:
    """Persistent history with preflight schema migration and cached duplicates."""
    def __init__(self,batch_size:int):
        backup_and_migrate_database()
        self.connection=sqlite3.connect(DB_PATH,timeout=30)
        self.connection.execute("PRAGMA journal_mode=DELETE")
        self.connection.execute("PRAGMA synchronous=FULL")
        self.connection.execute("PRAGMA busy_timeout=30000")
        self.completed=set(self.connection.execute("SELECT store_id,entry_id FROM exports WHERE status='completed'"))
        self.batch_size=max(1,int(batch_size));self.pending=0
    def is_duplicate(self,store_id,entry_id):return (store_id,entry_id) in self.completed
    def success(self,store_id,entry_id,item_class,item_type,folder,destination,size):
        self.connection.execute("""INSERT INTO exports(
            store_id,entry_id,item_class,item_type,source_folder,destination,file_size,status,updated,error)
            VALUES(?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(store_id,entry_id) DO UPDATE SET
            item_class=excluded.item_class,item_type=excluded.item_type,
            source_folder=excluded.source_folder,destination=excluded.destination,
            file_size=excluded.file_size,status='completed',updated=excluded.updated,error=''""",
            (store_id,entry_id,item_class,item_type,folder,str(destination),size,"completed",datetime.now().isoformat(),""))
        self.completed.add((store_id,entry_id));self.pending+=1;self._maybe_commit()
    def failure(self,store_id,entry_id,item_class,item_type,folder,error):
        self.connection.execute("""INSERT INTO exports(
            store_id,entry_id,item_class,item_type,source_folder,status,updated,error)
            VALUES(?,?,?,?,?,'failed',?,?)
            ON CONFLICT(store_id,entry_id) DO UPDATE SET
            item_class=excluded.item_class,item_type=excluded.item_type,
            source_folder=excluded.source_folder,status='failed',updated=excluded.updated,error=excluded.error""",
            (store_id,entry_id,item_class,item_type,folder,datetime.now().isoformat(),error))
        self.pending+=1;self._maybe_commit()
    def _maybe_commit(self):
        if self.pending>=self.batch_size:self.connection.commit();self.pending=0
    def close(self):self.connection.commit();self.connection.close()
'''

def main():
    if SOURCE is None:raise FileNotFoundError('Put this script beside inbox_email_exporter_v1_5_1.py')
    text=SOURCE.read_text(encoding='utf-8')
    if 'VERSION = "1.5.1"' in text:text=text.replace('VERSION = "1.5.1"','VERSION = "1.5.2"',1)
    elif 'VERSION = "1.5.0"' in text:text=text.replace('VERSION = "1.5.0"','VERSION = "1.5.2"',1)
    start=text.index('class History:')
    end=text.index('\n@dataclass',start)
    text=text[:start]+HISTORY+text[end:]
    # Run migration during startup, before Outlook scanning or a 45-minute run.
    marker='def startup_health_check():'
    pos=text.index(marker)
    body_start=text.index('\n',pos)+1
    text=text[:body_start]+'    backup_and_migrate_database()\n'+text[body_start:]
    OUTPUT.write_text(text,encoding='utf-8')
    py_compile.compile(str(OUTPUT),doraise=True)
    print('Created',OUTPUT)

if __name__=='__main__':main()
