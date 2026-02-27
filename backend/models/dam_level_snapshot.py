from datetime import datetime

from models.physchem import db


class DamLevelSnapshot(db.Model):
    __tablename__ = 'dam_level_snapshots'

    id = db.Column(db.Integer, primary_key=True)
    slot_datetime = db.Column(db.DateTime, nullable=False, unique=True, index=True)
    target_hour = db.Column(db.String(20), nullable=False)
    dam_level = db.Column(db.Float, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)
