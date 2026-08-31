from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from database import db, Donation, DonationItem, Campaign, Expense
import os
import math
from werkzeug.utils import secure_filename
from datetime import datetime
import openpyxl
from io import BytesIO
import jdatetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'

# ========================= تنظیمات دیتابیس =========================
# برای استفاده از SQLite (پیش‌فرض برای تست)
# app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///charity.db'

# در صورت تمایل به استفاده از MySQL، خط زیر را جایگزین کنید و رمز خود را اصلاح کنید
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:13681370Hb%40@localhost:3306/charity_db'

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# ===================================================================

# ================== فیلتر مخصوص فرمت اعداد ==================
@app.template_filter('to')
def format_number(value):
    """تبدیل عدد به فرمت هزارتایی با اعداد فارسی"""
    if value is None:
        return '۰'
    try:
        num = int(value)
        # تبدیل به هزارتا با جداکننده
        formatted = f"{num:,}"
        # تبدیل جداکننده به فارسی
        formatted = formatted.replace(',', '٬')
        # تبدیل اعداد انگلیسی به فارسی
        persian_digits = '۰۱۲۳۴۵۶۷۸۹'
        result = ''.join(persian_digits[int(d)] if d.isdigit() else d for d in formatted)
        return result
    except (ValueError, TypeError):
        return str(value)
# ============================================================
@app.template_filter('shamsi')
def filter_shamsi(date):
    """فیلتر برای نمایش تاریخ شمسی در قالب"""
    return to_shamsi(date)

@app.template_filter('shamsi_date')
def filter_shamsi_date(date):
    """فیلتر برای نمایش تاریخ شمسی بدون ساعت"""
    return to_shamsi_date(date)

# تنظیمات آپلود فایل
UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

db.init_app(app)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ================== توابع کمکی ==================

def get_balances(start_date=None, end_date=None):
    """محاسبه موجودی کل و موجودی به تفکیک هر دسته با قابلیت فیلتر تاریخ"""
    
    # فیلتر پایه برای واریزی‌های تأییدشده
    donation_query = Donation.query.filter(Donation.status == 'approved')
    if start_date:
        donation_query = donation_query.filter(Donation.created_at >= start_date)
    if end_date:
        donation_query = donation_query.filter(Donation.created_at <= end_date)
    
    # محاسبه مجموع واریزی‌ها
    total_donations = db.session.query(db.func.sum(DonationItem.amount)).join(
        Donation
    ).filter(
        Donation.status == 'approved'
    )
    if start_date:
        total_donations = total_donations.filter(Donation.created_at >= start_date)
    if end_date:
        total_donations = total_donations.filter(Donation.created_at <= end_date)
    total_donations = total_donations.scalar() or 0
    
    # فیلتر برای خرج‌ها
    expense_query = Expense.query
    if start_date:
        expense_query = expense_query.filter(Expense.created_at >= start_date)
    if end_date:
        expense_query = expense_query.filter(Expense.created_at <= end_date)
    total_expenses = db.session.query(db.func.sum(Expense.amount))
    if start_date:
        total_expenses = total_expenses.filter(Expense.created_at >= start_date)
    if end_date:
        total_expenses = total_expenses.filter(Expense.created_at <= end_date)
    total_expenses = total_expenses.scalar() or 0
    
    total_balance = total_donations - total_expenses

    # دسته‌های ثابت
    categories = ['صدقه', 'خیرات', 'زکات', 'نذر', 'کمک عمومی']
    balance_by_category = {}

    for cat in categories:
        donation_sum = db.session.query(db.func.sum(DonationItem.amount)).join(
            Donation
        ).filter(
            Donation.status == 'approved',
            DonationItem.category == cat
        )
        if start_date:
            donation_sum = donation_sum.filter(Donation.created_at >= start_date)
        if end_date:
            donation_sum = donation_sum.filter(Donation.created_at <= end_date)
        donation_sum = donation_sum.scalar() or 0
        
        expense_sum = db.session.query(db.func.sum(Expense.amount)).filter(
            Expense.category == cat
        )
        if start_date:
            expense_sum = expense_sum.filter(Expense.created_at >= start_date)
        if end_date:
            expense_sum = expense_sum.filter(Expense.created_at <= end_date)
        expense_sum = expense_sum.scalar() or 0
        
        balance_by_category[cat] = donation_sum - expense_sum

    # کمپین‌های فعال
    campaigns = Campaign.query.filter_by(is_active=True).all()
    for camp in campaigns:
        donation_sum = db.session.query(db.func.sum(DonationItem.amount)).join(
            Donation
        ).filter(
            Donation.status == 'approved',
            DonationItem.campaign_id == camp.id
        )
        if start_date:
            donation_sum = donation_sum.filter(Donation.created_at >= start_date)
        if end_date:
            donation_sum = donation_sum.filter(Donation.created_at <= end_date)
        donation_sum = donation_sum.scalar() or 0
        
        expense_sum = db.session.query(db.func.sum(Expense.amount)).filter(
            Expense.campaign_id == camp.id
        )
        if start_date:
            expense_sum = expense_sum.filter(Expense.created_at >= start_date)
        if end_date:
            expense_sum = expense_sum.filter(Expense.created_at <= end_date)
        expense_sum = expense_sum.scalar() or 0
        
        balance_by_category[camp.title] = donation_sum - expense_sum

    return total_balance, balance_by_category

def get_transactions(page=1, per_page=15, start_date=None, end_date=None):
    """دریافت لیست ترکیبی از واریزی‌ها و خرج‌ها با صفحه‌بندی و فیلتر تاریخ"""
    
    # فیلتر واریزی‌ها
    donation_query = Donation.query
    if start_date:
        donation_query = donation_query.filter(Donation.created_at >= start_date)
    if end_date:
        donation_query = donation_query.filter(Donation.created_at <= end_date)
    donations = donation_query.all()
    
    donation_items = []
    for d in donations:
        for item in d.items:
            cat_name = item.category if item.category else (item.campaign.title if item.campaign else '')
            donation_items.append({
                'id': f'd{d.id}',
                'real_id': d.id,
                'amount': item.amount,
                'category': cat_name,
                'created_at': d.created_at,
                'type': 'واریز',
                'name': d.donator_name,
                'description': '',
                'obj_type': 'donation_item',
                'status': d.status
            })

    # فیلتر خرج‌ها
    expense_query = Expense.query
    if start_date:
        expense_query = expense_query.filter(Expense.created_at >= start_date)
    if end_date:
        expense_query = expense_query.filter(Expense.created_at <= end_date)
    expenses = expense_query.all()
    
    expense_items = []
    for e in expenses:
        cat_name = e.category if e.category else (e.campaign.title if e.campaign else '')
        expense_items.append({
            'id': f'e{e.id}',
            'real_id': e.id,
            'amount': e.amount,
            'category': cat_name,
            'created_at': e.created_at,
            'type': 'خرج',
            'name': 'مدیر',
            'description': e.description,
            'obj_type': 'expense',
            'status': 'approved'
        })

    # ترکیب و مرتب‌سازی
    all_trans = donation_items + expense_items
    all_trans.sort(key=lambda x: x['created_at'], reverse=True)

    # صفحه‌بندی
    total = len(all_trans)
    total_pages = math.ceil(total / per_page) if total > 0 else 1
    start = (page - 1) * per_page
    end = start + per_page
    paginated = all_trans[start:end]

    return paginated, total, total_pages, page

# ================== توابع تبدیل تاریخ ==================
def to_shamsi(date):
    """تبدیل تاریخ میلادی به شمسی با فرمت کامل"""
    if date is None:
        return ''
    try:
        shamsi = jdatetime.datetime.fromgregorian(datetime=date)
        return shamsi.strftime('%Y/%m/%d %H:%M')
    except:
        return str(date)

def to_shamsi_date(date):
    """تبدیل تاریخ میلادی به شمسی فقط با فرمت تاریخ (بدون ساعت)"""
    if date is None:
        return ''
    try:
        shamsi = jdatetime.datetime.fromgregorian(datetime=date)
        return shamsi.strftime('%Y/%m/%d')
    except:
        return str(date)

# ========================================================
# ================== صفحه اصلی (فرم ثبت کمک) ==================
@app.route('/', methods=['GET', 'POST'])
def index():
    active_campaigns = Campaign.query.filter_by(is_active=True).all()

    if request.method == 'POST':
        name = request.form.get('name')
        categories = request.form.getlist('category[]')
        campaign_ids = request.form.getlist('campaign_id[]')
        amounts = request.form.getlist('amount[]')
        file = request.files.get('receipt')

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
            return redirect(url_for('index'))

        filename = None
        if file and allowed_file(file.filename):
            filename = secure_filename(f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        else:
            flash('فرمت فایل مجاز نیست (فقط عکس).', 'danger')
            return redirect(url_for('index'))

        donation = Donation(
            donator_name=name if name else 'ناشناس',
            receipt_image=filename,
            status='pending'
        )
        db.session.add(donation)
        db.session.flush()

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

# ================== پنل مدیریت (داشبورد) ==================
@app.route('/admin')
def admin_redirect():
    return redirect(url_for('dashboard'))

@app.route('/admin/dashboard')
def dashboard():
    page = request.args.get('page', 1, type=int)
    
    # دریافت فیلتر تاریخ از کوئری استرینگ
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    
    start_date = None
    end_date = None
    
    if start_date_str:
        try:
            # تبدیل تاریخ شمسی به میلادی
            shamsi_parts = start_date_str.split('/')
            if len(shamsi_parts) == 3:
                year, month, day = map(int, shamsi_parts)
                greg = jdatetime.date(year, month, day).togregorian()
                start_date = datetime(greg.year, greg.month, greg.day, 0, 0, 0)
        except:
            pass
    
    if end_date_str:
        try:
            shamsi_parts = end_date_str.split('/')
            if len(shamsi_parts) == 3:
                year, month, day = map(int, shamsi_parts)
                greg = jdatetime.date(year, month, day).togregorian()
                end_date = datetime(greg.year, greg.month, greg.day, 23, 59, 59)
        except:
            pass
    
    transactions, total, total_pages, current_page = get_transactions(page, 15, start_date, end_date)
    total_balance, balance_by_category = get_balances(start_date, end_date)

    # آمار با فیلتر تاریخ
    donation_count_query = Donation.query.filter_by(status='approved')
    expense_count_query = Expense.query
    pending_count_query = Donation.query.filter_by(status='pending')
    
    if start_date:
        donation_count_query = donation_count_query.filter(Donation.created_at >= start_date)
        expense_count_query = expense_count_query.filter(Expense.created_at >= start_date)
        pending_count_query = pending_count_query.filter(Donation.created_at >= start_date)
    if end_date:
        donation_count_query = donation_count_query.filter(Donation.created_at <= end_date)
        expense_count_query = expense_count_query.filter(Expense.created_at <= end_date)
        pending_count_query = pending_count_query.filter(Donation.created_at <= end_date)
    
    total_donations = donation_count_query.count()
    total_expenses = expense_count_query.count()
    pending_count = pending_count_query.count()

    # مقدار شروع و پایان برای نمایش در فرم
    start_date_display = start_date_str or ''
    end_date_display = end_date_str or ''

    return render_template('dashboard.html',
                           transactions=transactions,
                           total=total,
                           total_pages=total_pages,
                           current_page=current_page,
                           total_balance=total_balance,
                           balance_by_category=balance_by_category,
                           total_donations=total_donations,
                           total_expenses=total_expenses,
                           pending_count=pending_count,
                           start_date=start_date_display,
                           end_date=end_date_display)

# ================== تأیید/رد واریزی از داشبورد ==================
@app.route('/admin/approve/<int:donation_id>/<action>')
def approve_reject_donation(donation_id, action):
    donation = Donation.query.get_or_404(donation_id)
    if action == 'approve':
        donation.status = 'approved'
        flash(f'واریزی با شناسه {donation.id} تأیید شد.', 'success')
    elif action == 'reject':
        donation.status = 'rejected'
        flash(f'واریزی با شناسه {donation.id} رد شد.', 'success')
    else:
        flash('عملیات نامعتبر.', 'danger')
        return redirect(url_for('dashboard'))
    db.session.commit()
    return redirect(url_for('dashboard'))

# ================== مدیریت واریزی‌ها (ثبت دستی توسط مدیر) ==================
@app.route('/admin/donation/add', methods=['GET', 'POST'])
def add_donation():
    campaigns = Campaign.query.filter_by(is_active=True).all()

    if request.method == 'POST':
        name = request.form.get('name')
        categories = request.form.getlist('category[]')
        campaign_ids = request.form.getlist('campaign_id[]')
        amounts = request.form.getlist('amount[]')
        status = request.form.get('status', 'approved')
        file = request.files.get('receipt')

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
            return redirect(url_for('add_donation'))

        filename = None
        if file and allowed_file(file.filename):
            filename = secure_filename(f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        donation = Donation(
            donator_name=name if name else 'ناشناس',
            receipt_image=filename,
            status=status
        )
        db.session.add(donation)
        db.session.flush()

        for item in valid_items:
            donation_item = DonationItem(
                donation_id=donation.id,
                category=item['category'],
                campaign_id=item['campaign_id'],
                amount=item['amount']
            )
            db.session.add(donation_item)

        db.session.commit()
        flash('واریزی با موفقیت ثبت شد.', 'success')
        return redirect(url_for('dashboard'))

    return render_template('donation_form.html', campaigns=campaigns, action='افزودن', donation=None)

@app.route('/admin/donation/edit/<int:donation_id>', methods=['GET', 'POST'])
def edit_donation(donation_id):
    donation = Donation.query.get_or_404(donation_id)
    campaigns = Campaign.query.filter_by(is_active=True).all()

    if request.method == 'POST':
        name = request.form.get('name')
        categories = request.form.getlist('category[]')
        campaign_ids = request.form.getlist('campaign_id[]')
        amounts = request.form.getlist('amount[]')
        status = request.form.get('status', 'approved')

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

        donation.donator_name = name if name else 'ناشناس'
        donation.status = status

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
        flash('واریزی با موفقیت ویرایش شد.', 'success')
        return redirect(url_for('dashboard'))

    return render_template('donation_form.html', campaigns=campaigns, action='ویرایش', donation=donation)

@app.route('/admin/donation/delete/<int:donation_id>')
def delete_donation(donation_id):
    donation = Donation.query.get_or_404(donation_id)
    db.session.delete(donation)
    db.session.commit()
    flash('واریزی با موفقیت حذف شد.', 'success')
    return redirect(url_for('dashboard'))

# ================== مدیریت خرج‌ها ==================
@app.route('/admin/expense/add', methods=['GET', 'POST'])
def add_expense():
    campaigns = Campaign.query.filter_by(is_active=True).all()

    if request.method == 'POST':
        description = request.form.get('description')
        amount = request.form.get('amount')
        category = request.form.get('category')
        campaign_id = request.form.get('campaign_id')
        file = request.files.get('receipt')

        if not description or not amount or not amount.isdigit() or int(amount) <= 0:
            flash('توضیحات و مبلغ معتبر وارد کنید.', 'danger')
            return redirect(url_for('add_expense'))

        filename = None
        if file and allowed_file(file.filename):
            filename = secure_filename(f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        expense = Expense(
            description=description,
            amount=int(amount),
            category=category if category else None,
            campaign_id=int(campaign_id) if campaign_id else None,
            receipt_image=filename
        )
        db.session.add(expense)
        db.session.commit()

        flash('خرج با موفقیت ثبت شد.', 'success')
        return redirect(url_for('dashboard'))

    return render_template('expense_form.html', campaigns=campaigns, action='افزودن', expense=None)

@app.route('/admin/expense/edit/<int:expense_id>', methods=['GET', 'POST'])
def edit_expense(expense_id):
    expense = Expense.query.get_or_404(expense_id)
    campaigns = Campaign.query.filter_by(is_active=True).all()

    if request.method == 'POST':
        description = request.form.get('description')
        amount = request.form.get('amount')
        category = request.form.get('category')
        campaign_id = request.form.get('campaign_id')
        file = request.files.get('receipt')

        if not description or not amount or not amount.isdigit() or int(amount) <= 0:
            flash('توضیحات و مبلغ معتبر وارد کنید.', 'danger')
            return redirect(url_for('edit_expense', expense_id=expense_id))

        expense.description = description
        expense.amount = int(amount)
        expense.category = category if category else None
        expense.campaign_id = int(campaign_id) if campaign_id else None

        if file and allowed_file(file.filename):
            filename = secure_filename(f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            expense.receipt_image = filename

        db.session.commit()
        flash('خرج با موفقیت ویرایش شد.', 'success')
        return redirect(url_for('dashboard'))

    return render_template('expense_form.html', campaigns=campaigns, action='ویرایش', expense=expense)

@app.route('/admin/expense/delete/<int:expense_id>')
def delete_expense(expense_id):
    expense = Expense.query.get_or_404(expense_id)
    db.session.delete(expense)
    db.session.commit()
    flash('خرج با موفقیت حذف شد.', 'success')
    return redirect(url_for('dashboard'))

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

# ================== خروجی اکسل ==================
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
            label = item.category if item.category else (item.campaign.title if item.campaign else 'نامشخص')
            ws.append([idx, d.donator_name, label, item.amount, d.created_at.strftime('%Y-%m-%d')])
            idx += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name='گزارش_کمک‌ها.xlsx')

@app.route('/test')
def test_route():
    return "مسیر تست کار می‌کند!"

# ========================= اجرا =========================
with app.app_context():
    db.create_all()
    print("✅ جداول دیتابیس با موفقیت ساخته شدند.")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)