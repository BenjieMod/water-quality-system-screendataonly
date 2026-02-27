from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class PhysChemAnalysis(db.Model):
    __tablename__ = 'physchem_analysis'
    
    id = db.Column(db.Integer, primary_key=True)
    # General Info
    client = db.Column(db.String(200), nullable=False)
    source = db.Column(db.String(100))
    location = db.Column(db.String(200))
    date_collected = db.Column(db.Date)
    date_analyzed = db.Column(db.Date)
    date_submitted = db.Column(db.Date)
    file_prefix = db.Column(db.String(20))
    file_number = db.Column(db.String(20))
    or_number = db.Column(db.String(50))
    collected_by = db.Column(db.String(100))
    analyst = db.Column(db.String(100))
    
    # Physical Parameters
    pH = db.Column(db.Float)
    turbidity = db.Column(db.Float)
    color = db.Column(db.Float)
    total_dissolved_solids = db.Column(db.Float)
    
    # Chemical Parameters
    iron = db.Column(db.Float)
    chloride = db.Column(db.Float)
    copper = db.Column(db.Float)
    chromium = db.Column(db.Float)
    manganese = db.Column(db.Float)
    total_hardness = db.Column(db.Float)
    sulfate = db.Column(db.Float)
    nitrate = db.Column(db.Float)
    nitrite = db.Column(db.Float)
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'client': self.client,
            'source': self.source,
            'location': self.location,
            'date_collected': self.date_collected.isoformat() if self.date_collected else None,
            'date_analyzed': self.date_analyzed.isoformat() if self.date_analyzed else None,
            'date_submitted': self.date_submitted.isoformat() if self.date_submitted else None,
            'file_prefix': self.file_prefix,
            'file_number': self.file_number,
            'or_number': self.or_number,
            'collected_by': self.collected_by,
            'analyst': self.analyst,
            'pH': self.pH,
            'turbidity': self.turbidity,
            'color': self.color,
            'total_dissolved_solids': self.total_dissolved_solids,
            'iron': self.iron,
            'chloride': self.chloride,
            'copper': self.copper,
            'chromium': self.chromium,
            'manganese': self.manganese,
            'total_hardness': self.total_hardness,
            'sulfate': self.sulfate,
            'nitrate': self.nitrate,
            'nitrite': self.nitrite,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }
