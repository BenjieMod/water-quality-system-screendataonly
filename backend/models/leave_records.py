from models.physchem import db
from datetime import datetime

class CTOApplication(db.Model):
    __tablename__ = 'cto_applications'
    
    id = db.Column(db.Integer, primary_key=True)
    
    # Employee Information
    employee_no = db.Column(db.String(50), nullable=False)
    employee_name = db.Column(db.String(200), nullable=False)
    date_filed = db.Column(db.Date, nullable=False)
    
    # CTO Details
    date_covered_description = db.Column(db.String(200))  # e.g., "Nov 1 to Nov 3"
    from_date = db.Column(db.Date, nullable=False)
    to_date = db.Column(db.Date, nullable=False)
    total_hours = db.Column(db.Float, nullable=False)
    
    # Signatures
    applicant_signature = db.Column(db.Text)  # Base64 signature
    recommending_approval_name = db.Column(db.String(200))
    recommending_approval_title = db.Column(db.String(100))  # e.g., "OIC-WQD"
    recommending_signature = db.Column(db.Text)  # Base64 signature
    
    # Status
    status = db.Column(db.String(50), default='Pending')  # Pending, Approved, Rejected
    
    # Generated file
    excel_file = db.Column(db.String(500))
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'employee_no': self.employee_no,
            'employee_name': self.employee_name,
            'date_filed': self.date_filed.isoformat() if self.date_filed else None,
            'date_covered_description': self.date_covered_description,
            'from_date': self.from_date.isoformat() if self.from_date else None,
            'to_date': self.to_date.isoformat() if self.to_date else None,
            'total_hours': self.total_hours,
            'applicant_signature': self.applicant_signature,
            'recommending_approval_name': self.recommending_approval_name,
            'recommending_approval_title': self.recommending_approval_title,
            'recommending_signature': self.recommending_signature,
            'status': self.status,
            'excel_file': self.excel_file,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }


class LeaveApplication(db.Model):
    __tablename__ = 'leave_applications'
    
    id = db.Column(db.Integer, primary_key=True)
    
    # Employee Information
    employee_name = db.Column(db.String(200), nullable=False)
    date_filed = db.Column(db.Date, nullable=False)
    
    # Leave Type (store as JSON array of selected types)
    leave_types = db.Column(db.Text)  # JSON: ["Vacation Leave", "Sick Leave", etc.]
    other_leave_type = db.Column(db.String(200))  # If "Others" is selected
    
    # Leave Period
    from_date = db.Column(db.Date, nullable=False)
    to_date = db.Column(db.Date, nullable=False)
    day_off = db.Column(db.String(100))
    
    # Signatures
    applicant_signature = db.Column(db.Text)  # Base64 signature
    recommending_approval_name = db.Column(db.String(200))
    recommending_approval_title = db.Column(db.String(100))
    recommending_signature = db.Column(db.Text)  # Base64 signature
    date_signed = db.Column(db.Date)
    
    # Status
    status = db.Column(db.String(50), default='Pending')  # Pending, Approved, Rejected
    
    # Generated file
    excel_file = db.Column(db.String(500))
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        import json
        return {
            'id': self.id,
            'employee_name': self.employee_name,
            'date_filed': self.date_filed.isoformat() if self.date_filed else None,
            'leave_types': json.loads(self.leave_types) if self.leave_types else [],
            'other_leave_type': self.other_leave_type,
            'from_date': self.from_date.isoformat() if self.from_date else None,
            'to_date': self.to_date.isoformat() if self.to_date else None,
            'day_off': self.day_off,
            'applicant_signature': self.applicant_signature,
            'recommending_approval_name': self.recommending_approval_name,
            'recommending_approval_title': self.recommending_approval_title,
            'recommending_signature': self.recommending_signature,
            'date_signed': self.date_signed.isoformat() if self.date_signed else None,
            'status': self.status,
            'excel_file': self.excel_file,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }


class LeaveCredits(db.Model):
    __tablename__ = 'leave_credits'
    
    id = db.Column(db.Integer, primary_key=True)
    
    # Employee Information
    employee_no = db.Column(db.String(50), unique=True, nullable=False)
    employee_name = db.Column(db.String(200), nullable=False)
    
    # Leave Balances (in days)
    vacation_leave = db.Column(db.Float, default=0)
    sick_leave = db.Column(db.Float, default=0)
    cto_hours = db.Column(db.Float, default=0)  # CTO in hours
    
    # Metadata
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'employee_no': self.employee_no,
            'employee_name': self.employee_name,
            'vacation_leave': self.vacation_leave,
            'sick_leave': self.sick_leave,
            'cto_hours': self.cto_hours,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }
    

class Employee(db.Model):
    __tablename__ = 'employees'
    
    id = db.Column(db.Integer, primary_key=True)
    employee_no = db.Column(db.String(50), unique=True, nullable=False)
    employee_name = db.Column(db.String(200), nullable=False)
    signature = db.Column(db.Text)  # Base64 signature image
    position = db.Column(db.String(200))
    department = db.Column(db.String(200))
    
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'employee_no': self.employee_no,
            'employee_name': self.employee_name,
            'signature': self.signature,
            'position': self.position,
            'department': self.department
        }