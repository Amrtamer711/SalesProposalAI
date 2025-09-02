import sqlite3
from pathlib import Path
from datetime import datetime
from typing import Optional

DB_PATH = Path(__file__).parent / "proposals.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS proposals_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    submitted_by TEXT NOT NULL,
    client_name TEXT NOT NULL,
    date_generated TEXT NOT NULL,
    package_type TEXT NOT NULL,
    locations TEXT NOT NULL,
    total_amount TEXT NOT NULL
);
"""


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=5.0, isolation_level=None)
    # Enable WAL and set busy timeout
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout=5000;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


def init_db() -> None:
    conn = _connect()
    try:
        conn.execute("BEGIN")
        conn.execute(SCHEMA)
        conn.execute("COMMIT")
    finally:
        conn.close()


def log_proposal(
    submitted_by: str,
    client_name: str,
    package_type: str,
    locations: str,
    total_amount: str,
    date_generated: Optional[str] = None,
) -> None:
    if not date_generated:
        date_generated = datetime.now().isoformat()

    conn = _connect()
    try:
        conn.execute("BEGIN")
        conn.execute(
            """
            INSERT INTO proposals_log (submitted_by, client_name, date_generated, package_type, locations, total_amount)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (submitted_by, client_name, date_generated, package_type, locations, total_amount),
        )
        conn.execute("COMMIT")
    finally:
        conn.close()


def export_to_excel() -> str:
    """Export proposals log to Excel file and return the file path."""
    import pandas as pd
    import tempfile
    from datetime import datetime
    
    conn = _connect()
    try:
        # Read all proposals into a DataFrame
        df = pd.read_sql_query(
            "SELECT * FROM proposals_log ORDER BY date_generated DESC",
            conn
        )
        
        # Convert date_generated to datetime for better Excel formatting
        df['date_generated'] = pd.to_datetime(df['date_generated'])
        
        # Create a temporary Excel file
        temp_file = tempfile.NamedTemporaryFile(
            delete=False, 
            suffix=f'_proposals_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        temp_file.close()
        
        # Write to Excel with formatting
        with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Proposals', index=False)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Proposals']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Add filters
            worksheet.auto_filter.ref = worksheet.dimensions
        
        return temp_file.name
        
    finally:
        conn.close()


def get_proposals_summary() -> dict:
    """Get a summary of proposals for display."""
    conn = _connect()
    try:
        cursor = conn.cursor()
        
        # Total count
        cursor.execute("SELECT COUNT(*) FROM proposals_log")
        total_count = cursor.fetchone()[0]
        
        # Count by package type
        cursor.execute("""
            SELECT package_type, COUNT(*) 
            FROM proposals_log 
            GROUP BY package_type
        """)
        by_type = dict(cursor.fetchall())
        
        # Recent proposals
        cursor.execute("""
            SELECT client_name, locations, date_generated 
            FROM proposals_log 
            ORDER BY date_generated DESC 
            LIMIT 5
        """)
        recent = cursor.fetchall()
        
        return {
            "total_proposals": total_count,
            "by_package_type": by_type,
            "recent_proposals": [
                {
                    "client": row[0],
                    "locations": row[1],
                    "date": row[2]
                }
                for row in recent
            ]
        }
        
    finally:
        conn.close()


# Initialize DB on import
init_db() 