from datetime import datetime

from models.physchem import db


class TurbiditySnapshot(db.Model):
    __tablename__ = 'turbidity_snapshots'

    id = db.Column(db.Integer, primary_key=True)
    slot_datetime = db.Column(db.DateTime, nullable=False, unique=True, index=True)
    target_hour = db.Column(db.String(20), nullable=False)
    turbidity = db.Column(db.Float, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)
