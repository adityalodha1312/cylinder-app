from db import db
from datetime import datetime

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(50), default='driver')
    name = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'Username': self.username,
            'Password': self.password,
            'Role': self.role,
            'Name': self.name or self.username
        }

class Customer(db.Model):
    __tablename__ = 'customers'
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.String(20), unique=True)
    name = db.Column(db.String(255), nullable=False)
    email = db.Column(db.String(255))
    phone = db.Column(db.String(50))
    address = db.Column(db.Text)
    cold_call_done = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'Customer ID': self.customer_id,
            'Name': self.name,
            'Email': self.email or '',
            'Phone': self.phone or '',
            'Address': self.address or '',
            'Cold Call Done': self.cold_call_done
        }

class Cylinder(db.Model):
    __tablename__ = 'cylinders'
    id = db.Column(db.Integer, primary_key=True)
    uid = db.Column(db.String(100), unique=True, nullable=False)
    gas_type = db.Column(db.String(50))
    cylinder_type = db.Column(db.String(50))
    owner = db.Column(db.String(100), default='Depot')
    status = db.Column(db.String(20), default='Active')
    location = db.Column(db.String(255), default='Depot')
    last_activity_date = db.Column(db.String(50)) # Keep as string to match sheet format 'dd-mm-yyyy'
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'uid': self.uid,
            'gas_type': self.gas_type or '',
            'cylinder_type': self.cylinder_type or '',
            'owner': self.owner or '',
            'status': self.status or 'Active',
            'location': self.location or 'Depot',
            'last_activity': self.last_activity_date or ''
        }

class CylinderMaintenance(db.Model):
    __tablename__ = 'cylinder_maintenance'
    id = db.Column(db.Integer, primary_key=True)
    cylinder_uid = db.Column(db.String(100), unique=True, nullable=False)
    water_capacity = db.Column(db.String(50))
    fill_pressure = db.Column(db.String(50))
    gas_capacity = db.Column(db.String(50))
    unit = db.Column(db.String(20))
    is_mixture = db.Column(db.String(20), default='No')
    mix_ratio = db.Column(db.String(100))
    manufacture_date = db.Column(db.String(50))
    last_hydro_date = db.Column(db.String(50))
    next_hydro_due = db.Column(db.String(50))
    hydro_test_status = db.Column(db.String(50))
    cert_no = db.Column(db.String(100))
    is_uhp = db.Column(db.String(20), default='No')

    def to_dict(self):
        return {
            'uid': self.cylinder_uid,
            'water_capacity': self.water_capacity or '',
            'fill_pressure': self.fill_pressure or '',
            'gas_capacity': self.gas_capacity or '',
            'unit': self.unit or '',
            'is_mixture': self.is_mixture or 'No',
            'mix_ratio': self.mix_ratio or '',
            'manufacture_date': self.manufacture_date or '',
            'last_hydro_date': self.last_hydro_date or '',
            'next_hydro_due': self.next_hydro_due or '',
            'hydro_test_status': self.hydro_test_status or '',
            'cert_no': self.cert_no or '',
            'is_uhp': self.is_uhp or 'No'
        }

class Scan(db.Model):
    __tablename__ = 'scans'
    id = db.Column(db.Integer, primary_key=True)
    scan_date = db.Column(db.String(50), nullable=False)
    scan_time = db.Column(db.String(50))
    driver = db.Column(db.String(100))
    action = db.Column(db.String(50), nullable=False)
    cylinder_uid = db.Column(db.String(100), nullable=False)
    customer = db.Column(db.String(255))
    gas_type = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'date': self.scan_date,
            'time': self.scan_time or '',
            'driver': self.driver or '',
            'action': self.action,
            'uid': self.cylinder_uid,
            'customer': self.customer or '',
            'gas_type': self.gas_type or ''
        }

class CustomerMap(db.Model):
    __tablename__ = 'customer_map'
    id = db.Column(db.Integer, primary_key=True)
    scan_date = db.Column(db.String(50))
    scan_time = db.Column(db.String(50))
    driver = db.Column(db.String(100))
    action = db.Column(db.String(50))
    count = db.Column(db.Integer, default=0)
    uids = db.Column(db.Text)
    customer = db.Column(db.String(255))
    send_receipt = db.Column(db.Boolean, default=False)
    receipt_status = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'date': self.scan_date,
            'time': self.scan_time or '',
            'driver': self.driver or '',
            'action': self.action or '',
            'count': self.count or 0,
            'uids': self.uids or '',
            'customer': self.customer or '',
            'send_receipt': self.send_receipt,
            'receipt_status': self.receipt_status or ''
        }

class BulkTank(db.Model):
    __tablename__ = 'bulk_tanks'
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.String(50), nullable=False)
    gas = db.Column(db.String(50), nullable=False)
    opening = db.Column(db.Float, default=0.0)
    dead_volume = db.Column(db.Float, default=0.0)
    capacity = db.Column(db.Float, default=0.0)
    unit = db.Column(db.String(20))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'date': self.date,
            'gas': self.gas,
            'opening': self.opening,
            'dead_volume': self.dead_volume,
            'capacity': self.capacity,
            'unit': self.unit or ''
        }

class Product(db.Model):
    __tablename__ = 'products'
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(100))
    gas_type = db.Column(db.String(50))
    cylinder_type = db.Column(db.String(50))
    gas_per_cyl = db.Column(db.Float, default=0.0)
    unit = db.Column(db.String(20))
    is_virtual = db.Column(db.Boolean, default=False)

    def to_dict(self):
        return {
            'id': self.product_id,
            'name': self.name or '',
            'gas_type': self.gas_type or '',
            'cylinder_type': self.cylinder_type or '',
            'gas_per_cyl': self.gas_per_cyl,
            'unit': self.unit or '',
            'is_virtual': self.is_virtual
        }

class DuraGasHistory(db.Model):
    __tablename__ = 'dura_gas_history'
    id                  = db.Column(db.Integer, primary_key=True)
    cylinder_uid        = db.Column(db.String(100), nullable=False)
    gas_filled          = db.Column(db.String(50), nullable=False)
    previous_gas        = db.Column(db.String(50))
    purge_required      = db.Column(db.Boolean, default=False)
    purge_acknowledged  = db.Column(db.Boolean, default=False)
    operator            = db.Column(db.String(100))
    fill_date           = db.Column(db.String(50))
    fill_time           = db.Column(db.String(50))
    created_at          = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'cylinder_uid': self.cylinder_uid,
            'gas_filled': self.gas_filled,
            'previous_gas': self.previous_gas,
            'purge_required': self.purge_required,
            'purge_acknowledged': self.purge_acknowledged,
            'operator': self.operator,
            'fill_date': self.fill_date,
            'fill_time': self.fill_time
        }

class SystemSetting(db.Model):
    __tablename__ = 'system_settings'
    key = db.Column(db.String(100), primary_key=True)
    value = db.Column(db.String(255), nullable=False)

class AdminScanLog(db.Model):
    __tablename__ = 'admin_scan_logs'
    id = db.Column(db.Integer, primary_key=True)
    scan_date = db.Column(db.String(50), nullable=False)
    scan_time = db.Column(db.String(50))
    cylinder_uid = db.Column(db.String(100), nullable=False)
    gas_type = db.Column(db.String(50))
    customer = db.Column(db.String(255))
    action = db.Column(db.String(50), nullable=False)
    admin_name = db.Column(db.String(100))
    last_known_customer = db.Column(db.String(255))
    last_activity_date = db.Column(db.String(50))
    days_outstanding = db.Column(db.Integer)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'date': self.scan_date,
            'time': self.scan_time or '',
            'uid': self.cylinder_uid,
            'gas_type': self.gas_type or '',
            'customer': self.customer or '',
            'action': self.action,
            'admin_name': self.admin_name or '',
            'last_known_customer': self.last_known_customer or '',
            'last_activity_date': self.last_activity_date or '',
            'days_outstanding': self.days_outstanding
        }

class AccountsBatch(db.Model):
    __tablename__ = 'accounts_batches'
    id          = db.Column(db.Integer, primary_key=True)
    batch_ref   = db.Column(db.String(30), unique=True)
    batch_date  = db.Column(db.String(50), nullable=False)
    batch_time  = db.Column(db.String(50))
    customer    = db.Column(db.String(255))
    admin_name  = db.Column(db.String(100))
    status      = db.Column(db.String(20), default='Pending')
    billed_at   = db.Column(db.DateTime)
    billed_by   = db.Column(db.String(100))
    notes       = db.Column(db.Text)
    created_at  = db.Column(db.DateTime, default=datetime.utcnow)
    items       = db.relationship('AccountsBatchItem', backref='batch', lazy=True, cascade='all, delete-orphan')

    def gas_summary(self):
        """Returns dict of gas_type -> count for all items in batch."""
        summary = {}
        for item in self.items:
            g = item.gas_type or 'Unknown'
            summary[g] = summary.get(g, 0) + 1
        return summary

    def to_dict(self):
        return {
            'id': self.id,
            'batch_ref': self.batch_ref or '',
            'batch_date': self.batch_date,
            'batch_time': self.batch_time or '',
            'customer': self.customer or '',
            'admin_name': self.admin_name or '',
            'status': self.status or 'Pending',
            'billed_at': self.billed_at.strftime('%d-%m-%Y %H:%M') if self.billed_at else '',
            'billed_by': self.billed_by or '',
            'notes': self.notes or '',
            'gas_summary': self.gas_summary(),
            'total_cylinders': len(self.items)
        }

class AccountsBatchItem(db.Model):
    __tablename__ = 'accounts_batch_items'
    id           = db.Column(db.Integer, primary_key=True)
    batch_id     = db.Column(db.Integer, db.ForeignKey('accounts_batches.id'), nullable=False)
    cylinder_uid = db.Column(db.String(100))
    gas_type     = db.Column(db.String(50))

    def to_dict(self):
        return {
            'id': self.id,
            'batch_id': self.batch_id,
            'cylinder_uid': self.cylinder_uid or '',
            'gas_type': self.gas_type or ''
        }
