from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from database import db, Donation, DonationItem, Campaign
import os
from werkzeug.utils import secure_filename
from datetime import datetime
import openpyxl
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'

# ========================= تنظیمات اتصال به MySQL =========================
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///charity.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# =======================================================================

# تنظیمات آپلود فایل
UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

db.init_app(app)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ======================== صفحه اصلی (فرم ثبت کمک) ========================
@app.route('/', methods=['GET', 'POST'])
def index():
    # دریافت لیست کمپین‌های فعال
    active_campaigns = Campaign.query.filter_by(is_active=True).all()
    
    if request.method == 'POST':
        name = request.form.get('name')
        categories = request.form.getlist('category[]')  # گزینه‌های ثابت
        campaign_ids = request.form.getlist('campaign_id[]')  # شناسه کمپین‌های انتخاب‌شده
        amounts = request.form.getlist('amount[]')
        file = request.files.get('receipt')

        # اعتبارسنجی: حداقل یک ردیف معتبر
        valid_items = []
        for idx, amt in enumerate(amounts):
            if not amt or not amt.isdigit() or int(amt) <= 0:
                continue
            # مشخص کنید که این ردیف مربوط به گزینه ثابت است یا کمپین
            if idx < len(categories) and categories[idx]:
                # گزینه ثابت
                valid_items.append({
                    'category': categories[idx],
                    'campaign_id': None,
                    'amount': int(amt)
                })
            elif idx < len(campaign_ids) and campaign_ids[idx]:
                # کمپین
                valid_items.append({
                    'category': None,
                    'campaign_id': int(campaign_ids[idx]),
                    'amount': int(amt)
                })

        if not valid_items:
            flash('حداقل یک ردیف معتبر با مبلغ مثبت وارد کنید.', 'danger')
            return redirect(url_for('index'))

        # ذخیره فیش
        filename = None
        if file and allowed_file(file.filename):
            filename = secure_filename(f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        else:
            flash('فرمت فایل مجاز نیست (فقط عکس).', 'danger')
            return redirect(url_for('index'))

        # ایجاد رکورد اصلی
        donation = Donation(
            donator_name=name if name else 'ناشناس',
            receipt_image=filename,
            status='pending'
        )
        db.session.add(donation)
        db.session.flush()

        # ایجاد آیتم‌ها
        for item in valid_items:
            donation_item = DonationItem(
                donation_id=donation.id,
                category=item['category'],
                campaign_id=item['campaign_id'],
                amount=item['amount']
            )
            db.session.add(donation_item)

        db.session.commit()
        flash('کمک شما با موفقیت ثبت شد و در انتظار تأیید است.', 'success')
        return redirect(url_for('index'))

    return render_template('index.html', campaigns=active_campaigns)

# =========================== پنل مدیریت ===========================
@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if request.method == 'POST':
        donation_id = request.form.get('donation_id')
        action = request.form.get('action')
        donation = Donation.query.get(donation_id)
        if donation:
            if action == 'approve':
                donation.status = 'approved'
            elif action == 'reject':
                donation.status = 'rejected'
            db.session.commit()
            flash('وضعیت به‌روزرسانی شد.', 'success')
        return redirect(url_for('admin'))

    donations = Donation.query.order_by(Donation.created_at.desc()).all()
    total_pending = Donation.query.filter_by(status='pending').count()
    total_approved = Donation.query.filter_by(status='approved').count()
    total_amount = db.session.query(db.func.sum(DonationItem.amount)).join(Donation).filter(Donation.status == 'approved').scalar() or 0

    return render_template('admin.html',
                           donations=donations,
                           pending=total_pending,
                           approved=total_approved,
                           total_amount=total_amount)

# ================== مدیریت کمپین‌ها ==================
@app.route('/admin/campaigns')
def admin_campaigns():
    campaigns = Campaign.query.order_by(Campaign.created_at.desc()).all()
    return render_template('campaigns.html', campaigns=campaigns)

@app.route('/admin/campaigns/add', methods=['GET', 'POST'])
def add_campaign():
    if request.method == 'POST':
        title = request.form.get('title')
        description = request.form.get('description')
        target_amount = request.form.get('target_amount')
        is_active = request.form.get('is_active') == 'on'

        if not title:
            flash('عنوان الزامی است.', 'danger')
            return redirect(url_for('add_campaign'))

        campaign = Campaign(
            title=title,
            description=description,
            target_amount=int(target_amount) if target_amount and target_amount.isdigit() else None,
            is_active=is_active
        )
        db.session.add(campaign)
        db.session.commit()
        flash('کمپین با موفقیت ایجاد شد.', 'success')
        return redirect(url_for('admin_campaigns'))

    return render_template('campaign_form.html', action='افزودن', campaign=None)

@app.route('/admin/campaigns/edit/<int:campaign_id>', methods=['GET', 'POST'])
def edit_campaign(campaign_id):
    campaign = Campaign.query.get_or_404(campaign_id)
    if request.method == 'POST':
        campaign.title = request.form.get('title')
        campaign.description = request.form.get('description')
        target_amount = request.form.get('target_amount')
        campaign.target_amount = int(target_amount) if target_amount and target_amount.isdigit() else None
        campaign.is_active = request.form.get('is_active') == 'on'
        db.session.commit()
        flash('کمپین با موفقیت ویرایش شد.', 'success')
        return redirect(url_for('admin_campaigns'))

    return render_template('campaign_form.html', action='ویرایش', campaign=campaign)

@app.route('/admin/campaigns/delete/<int:campaign_id>')
def delete_campaign(campaign_id):
    campaign = Campaign.query.get_or_404(campaign_id)
    db.session.delete(campaign)
    db.session.commit()
    flash('کمپین حذف شد.', 'success')
    return redirect(url_for('admin_campaigns'))

# ================== ویرایش ثبت توسط مدیر ==================
@app.route('/admin/edit/<int:donation_id>', methods=['GET', 'POST'])
def edit_donation(donation_id):
    donation = Donation.query.get_or_404(donation_id)
    campaigns = Campaign.query.filter_by(is_active=True).all()

    if request.method == 'POST':
        categories = request.form.getlist('category[]')
        campaign_ids = request.form.getlist('campaign_id[]')
        amounts = request.form.getlist('amount[]')

        valid_items = []
        for idx, amt in enumerate(amounts):
            if not amt or not amt.isdigit() or int(amt) <= 0:
                continue
            if idx < len(categories) and categories[idx]:
                valid_items.append({
                    'category': categories[idx],
                    'campaign_id': None,
                    'amount': int(amt)
                })
            elif idx < len(campaign_ids) and campaign_ids[idx]:
                valid_items.append({
                    'category': None,
                    'campaign_id': int(campaign_ids[idx]),
                    'amount': int(amt)
                })

        if not valid_items:
            flash('حداقل یک ردیف معتبر با مبلغ مثبت وارد کنید.', 'danger')
            return redirect(url_for('edit_donation', donation_id=donation_id))

        # حذف آیتم‌های قدیمی
        DonationItem.query.filter_by(donation_id=donation.id).delete()
        for item in valid_items:
            donation_item = DonationItem(
                donation_id=donation.id,
                category=item['category'],
                campaign_id=item['campaign_id'],
                amount=item['amount']
            )
            db.session.add(donation_item)

        db.session.commit()
        flash('اطلاعات با موفقیت ویرایش شد.', 'success')
        return redirect(url_for('admin'))

    return render_template('edit.html', donation=donation, campaigns=campaigns)

# ========================= خروجی اکسل =========================
@app.route('/export/excel')
def export_excel():
    donations = Donation.query.filter_by(status='approved').all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "کمک‌های تأییدشده"
    ws.append(['ردیف', 'نام', 'دسته/کمپین', 'مبلغ (تومان)', 'تاریخ ثبت'])

    idx = 1
    for d in donations:
        for item in d.items:
            # تعیین عنوان
            if item.category:
                label = item.category
            elif item.campaign_id:
                label = item.campaign.title
            else:
                label = 'نامشخص'
            ws.append([idx, d.donator_name, label, item.amount, d.created_at.strftime('%Y-%m-%d')])
            idx += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name='گزارش_کمک‌ها.xlsx')

# ========================= اجرا =========================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
    
#if __name__ == '__main__':
 #   with app.app_context():
  #      db.create_all()
   # app.run(debug=True, host='0.0.0.0', port=5000)
