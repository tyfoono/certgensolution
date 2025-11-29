
import sqlite3
import os
from datetime import datetime
from typing import List, Dict, Optional


class CertificateDatabase:
    def __init__(self, db_path: str = "certificates.db"):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Initialize database with required tables"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Create certificates table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS certificates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                participant_name TEXT NOT NULL,
                participant_surname TEXT NOT NULL,
                participant_email TEXT NOT NULL,
                event_name TEXT NOT NULL,
                event_date TEXT NOT NULL,
                event_category TEXT NOT NULL,
                participant_role TEXT NOT NULL,
                participant_place TEXT,
                language TEXT NOT NULL
            )
        ''')

        # Create indexes for better performance
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_email ON certificates(participant_email)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_event ON certificates(event_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_date ON certificates(event_date)')

        conn.commit()
        conn.close()

    def add_certificate(self,
                        participant_name: str,
                        participant_surname: str,
                        participant_email: str,
                        event_name: str,
                        event_date: str,
                        event_category: str,
                        participant_role: str,
                        participant_place: str,
                        language: str) -> int:
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT INTO certificates (
                participant_name, participant_surname, participant_email,
                event_name, event_date, event_category, participant_role,
                participant_place, language
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            participant_name, participant_surname, participant_email,
            event_name, event_date, event_category, participant_role,
            participant_place, language
        ))

        certificate_id = cursor.lastrowid
        conn.commit()
        conn.close()

        return certificate_id

    def get_certificate(self, certificate_id: int) -> Optional[Dict]:
        """Get a specific certificate by ID"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM certificates WHERE id = ?', (certificate_id,))
        result = cursor.fetchone()
        conn.close()

        return dict(result) if result else None

