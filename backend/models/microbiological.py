from models.physchem import db
from datetime import datetime

class MicrobiologicalAnalysis(db.Model):
    __tablename__ = 'microbiological_analysis'
    
    id = db.Column(db.Integer, primary_key=True)
    
    # General Info
    client = db.Column(db.String(200), nullable=False)
    source = db.Column(db.String(100))
    location = db.Column(db.String(200))
    date_collected = db.Column(db.Date)
    date_analyzed = db.Column(db.Date)
    date_submitted = db.Column(db.Date)
    collected_by = db.Column(db.String(200))
    file_prefix = db.Column(db.String(20))
    file_number = db.Column(db.String(20))
    or_number = db.Column(db.String(50))
    
    # Generated files
    excel_file = db.Column(db.String(500))  # ADD THIS
    pdf_file = db.Column(db.String(500))  # ADD THIS
    
    # Microbiological Parameters
    total_coliform = db.Column(db.Float)
    e_coli = db.Column(db.Float)
    fecal_coliform = db.Column(db.Float)
    heterotrophic_plate_count = db.Column(db.Float)
    
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
        'collected_by': self.collected_by,
        'file_prefix': self.file_prefix,
        'file_number': self.file_number,
        'or_number': self.or_number,
        'excel_file': self.excel_file,  # ADD THIS
        'pdf_file': self.pdf_file,  # ADD THIS
        'total_coliform': self.total_coliform,
        'e_coli': self.e_coli,
        'fecal_coliform': self.fecal_coliform,
        'heterotrophic_plate_count': self.heterotrophic_plate_count,
        'created_at': self.created_at.isoformat() if self.created_at else None,
        'updated_at': self.updated_at.isoformat() if self.updated_at else None
    }