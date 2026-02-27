from models.physchem import db
from datetime import datetime

class WaterTreatmentReading(db.Model):
    __tablename__ = 'water_treatment_readings'
    
    id = db.Column(db.Integer, primary_key=True)
    
    # Reading date and time
    reading_datetime = db.Column(db.DateTime, nullable=False)
    
    # Measurements
    dam_level = db.Column(db.Float)
    raw_water_turbidity = db.Column(db.Float)  # NTU
    clarified_water_phase1 = db.Column(db.Float)  # NTU
    clarified_water_phase2 = db.Column(db.Float)  # NTU
    filtered_water_phase1 = db.Column(db.Float)  # NTU
    filtered_water_phase2 = db.Column(db.Float)  # NTU
    pac_dosage = db.Column(db.Float)  # L/min
    alum_dosage = db.Column(db.Float)  # % Pump
    
    # Notes field - NEW!
    notes = db.Column(db.Text)  # For documenting typhoons, river sources, observations
    
    # Generated files
    excel_file = db.Column(db.String(500))
    pdf_file = db.Column(db.String(500))
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'reading_datetime': self.reading_datetime.isoformat() if self.reading_datetime else None,
            'dam_level': self.dam_level,
            'raw_water_turbidity': self.raw_water_turbidity,
            'clarified_water_phase1': self.clarified_water_phase1,
            'clarified_water_phase2': self.clarified_water_phase2,
            'filtered_water_phase1': self.filtered_water_phase1,
            'filtered_water_phase2': self.filtered_water_phase2,
            'pac_dosage': self.pac_dosage,
            'alum_dosage': self.alum_dosage,
            'notes': self.notes,  # ADD THIS
            'excel_file': self.excel_file,
            'pdf_file': self.pdf_file,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }