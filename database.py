from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class Campaign(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text, nullable=True)
    target_amount = db.Column(db.Integer, nullable=True)  # مبلغ هدف (اختیاری)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items = db.relationship('DonationItem', backref='campaign', lazy=True)

class Donation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    donator_name = db.Column(db.String(100), nullable=True)
    receipt_image = db.Column(db.String(200), nullable=True)
    status = db.Column(db.String(20), default='pending')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items = db.relationship('DonationItem', backref='donation', lazy=True, cascade="all, delete-orphan")

class DonationItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    donation_id = db.Column(db.Integer, db.ForeignKey('donation.id'), nullable=False)
    category = db.Column(db.String(50), nullable=True)  # برای گزینه‌های ثابت (صدقه، خیرات و...)
    campaign_id = db.Column(db.Integer, db.ForeignKey('campaign.id'), nullable=True)  # برای کمپین‌ها
    amount = db.Column(db.Integer, nullable=False)
    
    # یکی از دو فیلد بالا باید مقدار داشته باشد