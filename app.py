from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from models import get_db, init_db
from excel_export import generate_excel, BUSINESS_TYPES
from io import BytesIO
import os

app = Flask(__name__)
app.secret_key = os.urandom(24)


@app.before_request
def before_request():
    init_db()


# ===== トップページ =====
@app.route('/')
def index():
    return render_template('index.html')


# ===== 顧客管理 =====
@app.route('/customers')
def customer_list():
    db = get_db()
    customers = db.execute('SELECT * FROM customers ORDER BY updated_at DESC').fetchall()
    db.close()
    return render_template('customers.html', customers=customers)


@app.route('/customers/new', methods=['GET', 'POST'])
def customer_new():
    if request.method == 'POST':
        db = get_db()
        db.execute('''
            INSERT INTO customers (name, name_kana, representative, representative_kana,
                corporate_number, capital_amount, corporation_type,
                postal_code, prefecture, city, address, phone, fax)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            request.form['name'],
            request.form.get('name_kana', ''),
            request.form.get('representative', ''),
            request.form.get('representative_kana', ''),
            request.form.get('corporate_number', ''),
            request.form.get('capital_amount', ''),
            request.form.get('corporation_type', 1),
            request.form.get('postal_code', ''),
            request.form.get('prefecture', ''),
            request.form.get('city', ''),
            request.form.get('address', ''),
            request.form.get('phone', ''),
            request.form.get('fax', ''),
        ))
        db.commit()
        db.close()
        flash('顧客を登録しました。', 'success')
        return redirect(url_for('customer_list'))
    return render_template('customer_form.html', customer=None)


@app.route('/customers/<int:customer_id>/edit', methods=['GET', 'POST'])
def customer_edit(customer_id):
    db = get_db()
    if request.method == 'POST':
        db.execute('''
            UPDATE customers SET name=?, name_kana=?, representative=?, representative_kana=?,
                corporate_number=?, capital_amount=?, corporation_type=?,
                postal_code=?, prefecture=?, city=?, address=?, phone=?, fax=?,
                updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        ''', (
            request.form['name'],
            request.form.get('name_kana', ''),
            request.form.get('representative', ''),
            request.form.get('representative_kana', ''),
            request.form.get('corporate_number', ''),
            request.form.get('capital_amount', ''),
            request.form.get('corporation_type', 1),
            request.form.get('postal_code', ''),
            request.form.get('prefecture', ''),
            request.form.get('city', ''),
            request.form.get('address', ''),
            request.form.get('phone', ''),
            request.form.get('fax', ''),
            customer_id,
        ))
        db.commit()
        db.close()
        flash('顧客情報を更新しました。', 'success')
        return redirect(url_for('customer_list'))
    customer = db.execute('SELECT * FROM customers WHERE id=?', (customer_id,)).fetchone()
    db.close()
    return render_template('customer_form.html', customer=customer)


@app.route('/customers/<int:customer_id>/delete', methods=['POST'])
def customer_delete(customer_id):
    db = get_db()
    db.execute('DELETE FROM applications WHERE customer_id=?', (customer_id,))
    db.execute('DELETE FROM customers WHERE id=?', (customer_id,))
    db.commit()
    db.close()
    flash('顧客を削除しました。', 'success')
    return redirect(url_for('customer_list'))


# ===== 申請書管理 =====
@app.route('/customers/<int:customer_id>/applications')
def application_list(customer_id):
    db = get_db()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (customer_id,)).fetchone()
    applications = db.execute(
        'SELECT * FROM applications WHERE customer_id=? ORDER BY updated_at DESC',
        (customer_id,)
    ).fetchall()
    db.close()
    return render_template('applications.html', customer=customer, applications=applications)


@app.route('/customers/<int:customer_id>/applications/new', methods=['GET', 'POST'])
def application_new(customer_id):
    db = get_db()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (customer_id,)).fetchone()
    if request.method == 'POST':
        business_types = ','.join(request.form.getlist('business_types'))
        existing_business_types = ','.join(request.form.getlist('existing_business_types'))
        db.execute('''
            INSERT INTO applications (customer_id, application_date, permit_type,
                governor_or_minister, permit_category, permit_number,
                permit_year, permit_month, permit_day,
                general_or_specific, application_category, validity_adjustment,
                side_business, side_business_type, permit_transfer_category,
                old_permit_number, old_permit_year, old_permit_month, old_permit_day,
                city_code, business_types, existing_business_types,
                applicant_name, applicant_address, proxy_name, proxy_address,
                contact_organization, contact_name, contact_phone, contact_fax)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            customer_id,
            request.form.get('application_date', ''),
            request.form.get('permit_type', ''),
            request.form.get('governor_or_minister', 2),
            request.form.get('permit_category', ''),
            request.form.get('permit_number', ''),
            request.form.get('permit_year', ''),
            request.form.get('permit_month', ''),
            request.form.get('permit_day', ''),
            request.form.get('general_or_specific', 1),
            request.form.get('application_category', 1),
            request.form.get('validity_adjustment', 2),
            request.form.get('side_business', 2),
            request.form.get('side_business_type', ''),
            request.form.get('permit_transfer_category', ''),
            request.form.get('old_permit_number', ''),
            request.form.get('old_permit_year', ''),
            request.form.get('old_permit_month', ''),
            request.form.get('old_permit_day', ''),
            request.form.get('city_code', ''),
            business_types,
            existing_business_types,
            request.form.get('applicant_name', ''),
            request.form.get('applicant_address', ''),
            request.form.get('proxy_name', ''),
            request.form.get('proxy_address', ''),
            request.form.get('contact_organization', ''),
            request.form.get('contact_name', ''),
            request.form.get('contact_phone', ''),
            request.form.get('contact_fax', ''),
        ))
        db.commit()
        db.close()
        flash('申請書を保存しました。', 'success')
        return redirect(url_for('application_list', customer_id=customer_id))
    db.close()
    return render_template('application_form.html', customer=customer, application=None,
                           business_types=BUSINESS_TYPES)


@app.route('/applications/<int:app_id>/edit', methods=['GET', 'POST'])
def application_edit(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (application['customer_id'],)).fetchone()
    if request.method == 'POST':
        business_types = ','.join(request.form.getlist('business_types'))
        existing_business_types = ','.join(request.form.getlist('existing_business_types'))
        db.execute('''
            UPDATE applications SET application_date=?, permit_type=?,
                governor_or_minister=?, permit_category=?, permit_number=?,
                permit_year=?, permit_month=?, permit_day=?,
                general_or_specific=?, application_category=?, validity_adjustment=?,
                side_business=?, side_business_type=?, permit_transfer_category=?,
                old_permit_number=?, old_permit_year=?, old_permit_month=?, old_permit_day=?,
                city_code=?, business_types=?, existing_business_types=?,
                applicant_name=?, applicant_address=?, proxy_name=?, proxy_address=?,
                contact_organization=?, contact_name=?, contact_phone=?, contact_fax=?,
                updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        ''', (
            request.form.get('application_date', ''),
            request.form.get('permit_type', ''),
            request.form.get('governor_or_minister', 2),
            request.form.get('permit_category', ''),
            request.form.get('permit_number', ''),
            request.form.get('permit_year', ''),
            request.form.get('permit_month', ''),
            request.form.get('permit_day', ''),
            request.form.get('general_or_specific', 1),
            request.form.get('application_category', 1),
            request.form.get('validity_adjustment', 2),
            request.form.get('side_business', 2),
            request.form.get('side_business_type', ''),
            request.form.get('permit_transfer_category', ''),
            request.form.get('old_permit_number', ''),
            request.form.get('old_permit_year', ''),
            request.form.get('old_permit_month', ''),
            request.form.get('old_permit_day', ''),
            request.form.get('city_code', ''),
            business_types,
            existing_business_types,
            request.form.get('applicant_name', ''),
            request.form.get('applicant_address', ''),
            request.form.get('proxy_name', ''),
            request.form.get('proxy_address', ''),
            request.form.get('contact_organization', ''),
            request.form.get('contact_name', ''),
            request.form.get('contact_phone', ''),
            request.form.get('contact_fax', ''),
            app_id,
        ))
        db.commit()
        db.close()
        flash('申請書を更新しました。', 'success')
        return redirect(url_for('application_list', customer_id=application['customer_id']))
    db.close()
    return render_template('application_form.html', customer=customer, application=application,
                           business_types=BUSINESS_TYPES)


@app.route('/applications/<int:app_id>/delete', methods=['POST'])
def application_delete(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    customer_id = application['customer_id']
    db.execute('DELETE FROM applications WHERE id=?', (app_id,))
    db.commit()
    db.close()
    flash('申請書を削除しました。', 'success')
    return redirect(url_for('application_list', customer_id=customer_id))


@app.route('/applications/<int:app_id>/export')
def application_export(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (application['customer_id'],)).fetchone()
    db.close()

    excel_data = generate_excel(dict(application), dict(customer))
    customer_name = customer['name'] or '申請書'
    filename = f"建設業許可申請書_{customer_name}.xlsx"

    return send_file(
        BytesIO(excel_data),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
