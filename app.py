from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from models import get_db, init_db
from excel_export import generate_excel, generate_officers_excel, generate_offices_list_excel, BUSINESS_TYPES
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
            INSERT INTO customers (name, name_kana, representative, representative_title, representative_kana,
                corporate_number, capital_amount, corporation_type,
                postal_code, prefecture, city, address, phone, fax)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            request.form['name'],
            request.form.get('name_kana', ''),
            request.form.get('representative', ''),
            request.form.get('representative_title', ''),
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
            UPDATE customers SET name=?, name_kana=?, representative=?, representative_title=?, representative_kana=?,
                corporate_number=?, capital_amount=?, corporation_type=?,
                postal_code=?, prefecture=?, city=?, address=?, phone=?, fax=?,
                updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        ''', (
            request.form['name'],
            request.form.get('name_kana', ''),
            request.form.get('representative', ''),
            request.form.get('representative_title', ''),
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


# ===== 役員管理 =====
@app.route('/applications/<int:app_id>/officers', methods=['GET'])
def officers_list(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (application['customer_id'],)).fetchone()
    officers = db.execute(
        'SELECT * FROM officers WHERE application_id=? ORDER BY sort_order',
        (app_id,)
    ).fetchall()
    db.close()
    return render_template('officers.html', customer=customer, application=application,
                           officers=officers)


@app.route('/applications/<int:app_id>/officers/save', methods=['POST'])
def officers_save(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    # 既存データを削除して再登録
    db.execute('DELETE FROM officers WHERE application_id=?', (app_id,))
    total_rows = int(request.form.get('total_rows', 0))
    sort_order = 0
    for i in range(total_rows):
        if request.form.get(f'delete_{i}'):
            continue
        last_name = request.form.get(f'last_name_{i}', '').strip()
        first_name = request.form.get(f'first_name_{i}', '').strip()
        last_name_kana = request.form.get(f'last_name_kana_{i}', '').strip()
        first_name_kana = request.form.get(f'first_name_kana_{i}', '').strip()
        role = request.form.get(f'role_{i}', '').strip()
        full_or_part = request.form.get(f'full_or_part_{i}', '').strip()
        if last_name or first_name:
            db.execute('''
                INSERT INTO officers (application_id, last_name, first_name, last_name_kana, first_name_kana, role, full_or_part, sort_order)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (app_id, last_name, first_name, last_name_kana, first_name_kana, role, full_or_part, sort_order))
            sort_order += 1
    db.commit()
    db.close()
    flash('役員情報を保存しました。', 'success')
    return redirect(url_for('officers_list', app_id=app_id))


@app.route('/applications/<int:app_id>/officers/export')
def officers_export(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (application['customer_id'],)).fetchone()
    officers = db.execute(
        'SELECT * FROM officers WHERE application_id=? ORDER BY sort_order',
        (app_id,)
    ).fetchall()
    db.close()

    excel_data = generate_officers_excel(
        [dict(o) for o in officers],
        application['application_date']
    )
    customer_name = customer['name'] or '役員一覧'
    filename = f"役員一覧表_{customer_name}.xlsx"

    return send_file(
        BytesIO(excel_data),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


# ===== 営業所管理 =====
@app.route('/applications/<int:app_id>/offices', methods=['GET'])
def offices_list(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (application['customer_id'],)).fetchone()
    offices = db.execute(
        'SELECT * FROM offices WHERE application_id=? ORDER BY office_type, sort_order',
        (app_id,)
    ).fetchall()
    db.close()
    return render_template('offices_list.html', customer=customer, application=application,
                           offices=offices, business_types=BUSINESS_TYPES)


@app.route('/applications/<int:app_id>/offices/save', methods=['POST'])
def offices_save(app_id):
    db = get_db()
    db.execute('DELETE FROM offices WHERE application_id=?', (app_id,))

    total_rows = int(request.form.get('total_rows', 0))
    for i in range(total_rows):
        if request.form.get(f'delete_{i}'):
            continue
        office_type = int(request.form.get(f'office_type_{i}', 2))
        name = request.form.get(f'name_{i}', '').strip()
        name_kana = request.form.get(f'name_kana_{i}', '').strip()
        business_types = ','.join(request.form.getlist(f'business_types_{i}'))

        # 主たる営業所は名称・業種、従たる営業所はフル入力
        if office_type == 1:
            main_name = name if name else '本店'
            main_name_kana = name_kana if name_kana else 'ホンテン'
            db.execute('''
                INSERT INTO offices (application_id, office_type, name, name_kana,
                    business_types, sort_order)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (app_id, 1, main_name, main_name_kana, business_types, 0))
        else:
            city_code = request.form.get(f'city_code_{i}', '').strip()
            prefecture = request.form.get(f'prefecture_{i}', '').strip()
            city = request.form.get(f'city_{i}', '').strip()
            address = request.form.get(f'address_{i}', '').strip()
            postal_code = request.form.get(f'postal_code_{i}', '').strip()
            phone = request.form.get(f'phone_{i}', '').strip()
            if name or business_types:
                db.execute('''
                    INSERT INTO offices (application_id, office_type, name, name_kana,
                        city_code, prefecture, city, address, postal_code, phone,
                        business_types, sort_order)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (app_id, 2, name, name_kana, city_code, prefecture, city,
                      address, postal_code, phone, business_types, i))

    db.commit()
    db.close()
    flash('営業所情報を保存しました。', 'success')
    return redirect(url_for('offices_list', app_id=app_id))


@app.route('/applications/<int:app_id>/offices/export')
def offices_export(app_id):
    db = get_db()
    application = db.execute('SELECT * FROM applications WHERE id=?', (app_id,)).fetchone()
    customer = db.execute('SELECT * FROM customers WHERE id=?', (application['customer_id'],)).fetchone()
    offices = db.execute(
        'SELECT * FROM offices WHERE application_id=? ORDER BY office_type, sort_order',
        (app_id,)
    ).fetchall()
    db.close()

    excel_data = generate_offices_list_excel(
        dict(application), dict(customer), [dict(o) for o in offices]
    )
    customer_name = customer['name'] or '営業所一覧'
    filename = f"営業所一覧_{customer_name}.xlsx"

    return send_file(
        BytesIO(excel_data),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
