import os
from io import BytesIO
from openpyxl import load_workbook

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '許可申請書.xlsx')

# 建設業29業種コード
BUSINESS_TYPES = [
    ('土', 'civil_engineering'),
    ('建', 'architecture'),
    ('大', 'carpentry'),
    ('左', 'plastering'),
    ('と', 'masonry'),
    ('石', 'stone'),
    ('屋', 'roofing'),
    ('電', 'electrical'),
    ('管', 'plumbing'),
    ('タ', 'tiling'),
    ('鋼', 'steel'),
    ('筋', 'rebar'),
    ('舗', 'paving'),
    ('しゆ', 'dredging'),
    ('板', 'sheet_metal'),
    ('ガ', 'glass'),
    ('塗', 'painting'),
    ('防', 'waterproofing'),
    ('内', 'interior'),
    ('機', 'machinery'),
    ('絶', 'insulation'),
    ('通', 'telecom'),
    ('園', 'landscaping'),
    ('井', 'well'),
    ('具', 'fixtures'),
    ('水', 'water'),
    ('消', 'fire_protection'),
    ('清', 'cleaning'),
    ('解', 'demolition'),
]

# 業種に対応するExcelセル列（結合セルの開始列）
BUSINESS_TYPE_COLUMNS = [
    'AR', 'AV', 'AZ', 'BD', 'BH', 'BL', 'BP', 'BT', 'BX', 'CB',
    'CF', 'CJ', 'CN', 'CR', 'CV', 'CZ', 'DD', 'DH', 'DL', 'DP',
    'DT', 'DX', 'EB', 'EF', 'EJ', 'EN', 'ER', 'EV', 'EZ',
]


def _fill_cells(ws, cells, text):
    """文字列を1文字ずつ指定セルに書き込む"""
    if not text:
        return
    for i, char in enumerate(text):
        if i >= len(cells):
            break
        ws[cells[i]] = char


def _fill_digits(ws, cells, number_str):
    """数字文字列を右詰めで1桁ずつセルに書き込む"""
    if not number_str:
        return
    digits = str(number_str).strip()
    # 右詰め
    start = len(cells) - len(digits)
    for i, d in enumerate(digits):
        pos = start + i
        if 0 <= pos < len(cells):
            ws[cells[pos]] = d


def generate_excel(application, customer):
    """申請データからExcelファイルを生成してバイトデータを返す"""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb['様式第一号']

    # ============================================================
    # 申請日（行9: 令和__年__月__日）→ 行政庁側記入欄のため空欄
    # ============================================================

    # ============================================================
    # 項番01: 許可番号
    # ============================================================
    # 般/特
    if application.get('general_or_specific'):
        gs = int(application['general_or_specific'])
        ws['CH21'] = '般' if gs == 1 else '特'

    # 許可番号（6桁: CZ21, DD21, DH21, DL21, DP21, DT21）
    permit_number_cells = ['CZ21', 'DD21', 'DH21', 'DL21', 'DP21', 'DT21']
    if application.get('permit_number'):
        _fill_digits(ws, permit_number_cells, application['permit_number'])

    # 許可年月日（年2桁: EL21,EP21 / 月2桁: EW21,FA21 / 日2桁: FH21,FL21）
    if application.get('permit_year'):
        _fill_digits(ws, ['EL21', 'EP21'], application['permit_year'])
    if application.get('permit_month'):
        _fill_digits(ws, ['EW21', 'FA21'], application['permit_month'])
    if application.get('permit_day'):
        _fill_digits(ws, ['FH21', 'FL21'], application['permit_day'])

    # ============================================================
    # 項番02: 申請の区分
    # ============================================================
    if application.get('application_category'):
        ws['AL25'] = str(application['application_category'])

    # 許可の有効期間の調整（1:する 2:しない）
    if application.get('validity_adjustment'):
        ws['FE25'] = str(application['validity_adjustment'])

    # ============================================================
    # 項番03: 申請年月日（年2桁: AY28,BC28 / 月2桁: BJ28,BN28 / 日2桁: BU28,BY28）
    # ============================================================
    if application.get('application_date'):
        parts = application['application_date'].split('-')
        if len(parts) == 3:
            year = int(parts[0]) - 2018
            _fill_digits(ws, ['AY28', 'BC28'], str(year))
            _fill_digits(ws, ['BJ28', 'BN28'], str(int(parts[1])))
            _fill_digits(ws, ['BU28', 'BY28'], str(int(parts[2])))

    # ============================================================
    # 項番04: 許可を受けようとする建設業（行34）
    # ============================================================
    if application.get('business_types'):
        selected = application['business_types'].split(',')
        for i, (label, code) in enumerate(BUSINESS_TYPES):
            if code in selected:
                col = BUSINESS_TYPE_COLUMNS[i]
                ws[f'{col}34'] = '1'

    # ============================================================
    # 項番05: 既に許可を受けている建設業（行37）
    # ============================================================
    if application.get('existing_business_types'):
        selected = application['existing_business_types'].split(',')
        for i, (label, code) in enumerate(BUSINESS_TYPES):
            if code in selected:
                col = BUSINESS_TYPE_COLUMNS[i]
                ws[f'{col}37'] = '1'

    # ============================================================
    # 項番06: 商号又は名称のフリガナ（行40, 20文字分）
    # ============================================================
    name_kana_cells = [
        'AR40', 'AY40', 'BF40', 'BM40', 'BT40', 'CA40', 'CH40', 'CO40', 'CV40', 'DC40',
        'DJ40', 'DQ40', 'DX40', 'EE40', 'EL40', 'ES40', 'EZ40', 'FG40', 'FN40', 'FU40',
    ]
    name_kana = customer.get('name_kana') or ''
    if customer.get('corporation_type') and int(customer['corporation_type']) == 1:
        # 法人の場合、法人格のフリガナを除去
        corp_kana_list = [
            'カブシキガイシャ', 'ユウゲンガイシャ', 'ゴウドウガイシャ', 'ゴウシガイシャ',
            'ゴウメイガイシャ', 'イッパンシャダンホウジン', 'イッパンザイダンホウジン',
            'コウエキシャダンホウジン', 'コウエキザイダンホウジン',
            'トクテイヒエイリカツドウホウジン', 'シャカイフクシホウジン',
            'ガッコウホウジン', 'イリョウホウジン',
        ]
        for corp_kana in corp_kana_list:
            name_kana = name_kana.replace(corp_kana, '').strip()
    _fill_cells(ws, name_kana_cells, name_kana)

    # ============================================================
    # 項番07: 商号又は名称（行46, 20文字分）
    # ============================================================
    name_cells = [
        'AR46', 'AY46', 'BF46', 'BM46', 'BT46', 'CA46', 'CH46', 'CO46', 'CV46', 'DC46',
        'DJ46', 'DQ46', 'DX46', 'EE46', 'EL46', 'ES46', 'EZ46', 'FG46', 'FN46', 'FU46',
    ]
    name = customer.get('name') or ''
    if customer.get('corporation_type') and int(customer['corporation_type']) == 1:
        # 法人格を略称に置換
        corp_replacements = [
            ('株式会社', '（株）'),
            ('有限会社', '（有）'),
            ('合同会社', '（合）'),
            ('合資会社', '（資）'),
            ('合名会社', '（名）'),
            ('一般社団法人', '（一社）'),
            ('一般財団法人', '（一財）'),
            ('公益社団法人', '（公社）'),
            ('公益財団法人', '（公財）'),
            ('特定非営利活動法人', '（特非）'),
            ('社会福祉法人', '（福）'),
            ('学校法人', '（学）'),
            ('医療法人', '（医）'),
        ]
        for full, short in corp_replacements:
            name = name.replace(full, short)
    _fill_cells(ws, name_cells, name)

    # ============================================================
    # 項番08: 代表者又は個人の氏名のフリガナ（行52, 20文字分）
    # ============================================================
    rep_kana_cells = [
        'AR52', 'AY52', 'BF52', 'BM52', 'BT52', 'CA52', 'CH52', 'CO52', 'CV52', 'DC52',
        'DJ52', 'DQ52', 'DX52', 'EE52', 'EL52', 'ES52', 'EZ52', 'FG52', 'FN52', 'FU52',
    ]
    _fill_cells(ws, rep_kana_cells, customer.get('representative_kana'))

    # ============================================================
    # 項番09: 代表者又は個人の氏名（行55, 10文字分）
    # ============================================================
    rep_name_cells = [
        'AR55', 'AY55', 'BF55', 'BM55', 'BT55', 'CA55', 'CH55', 'CO55', 'CV55', 'DC55',
    ]
    _fill_cells(ws, rep_name_cells, customer.get('representative'))

    # ============================================================
    # 項番10: 市区町村コード（行58, 5桁: AR58,AV58,AZ58,BD58,BH58）
    # ============================================================
    city_code_cells = ['AR58', 'AV58', 'AZ58', 'BD58', 'BH58']
    if application.get('city_code'):
        _fill_digits(ws, city_code_cells, application['city_code'])

    # 都道府県名（CF58 結合セル）
    if customer.get('prefecture'):
        ws['CF58'] = customer['prefecture']

    # 市区町村名（EK58 結合セル）
    if customer.get('city'):
        ws['EK58'] = customer['city']

    # ============================================================
    # 項番11: 主たる営業所の所在地（行61, 20文字分）
    # ============================================================
    address_cells = [
        'AR61', 'AY61', 'BF61', 'BM61', 'BT61', 'CA61', 'CH61', 'CO61', 'CV61', 'DC61',
        'DJ61', 'DQ61', 'DX61', 'EE61', 'EL61', 'ES61', 'EZ61', 'FG61', 'FN61', 'FU61',
    ]
    _fill_cells(ws, address_cells, customer.get('address'))

    # ============================================================
    # 項番12: 郵便番号（3桁: AR67,AV67,AZ67 / 4桁: BH67,BL67,BP67,BT67）
    # ============================================================
    if customer.get('postal_code'):
        parts = customer['postal_code'].replace('ー', '-').replace('－', '-').split('-')
        if len(parts) == 2:
            _fill_cells(ws, ['AR67', 'AV67', 'AZ67'], parts[0])
            _fill_cells(ws, ['BH67', 'BL67', 'BP67', 'BT67'], parts[1])

    # 電話番号（DD67〜EZ67, 最大12文字）
    phone_cells = [
        'DD67', 'DH67', 'DL67', 'DP67', 'DT67', 'DX67',
        'EB67', 'EF67', 'EJ67', 'EN67', 'ER67', 'EV67', 'EZ67',
    ]
    if customer.get('phone'):
        _fill_cells(ws, phone_cells, customer['phone'].replace('-', '').replace('ー', ''))

    # FAX番号（BQ70 結合セル）
    if customer.get('fax'):
        ws['BQ70'] = customer['fax']
    else:
        ws['BQ70'] = 'なし'

    # ============================================================
    # 項番13: 法人又は個人の別
    # ============================================================
    if customer.get('corporation_type'):
        ws['AR75'] = str(customer['corporation_type'])

    # 資本金額（千円）（行75: BT75,BX75,CB75,CF75,CJ75,CN75,CR75,CV75,CZ75）
    capital_cells = ['BT75', 'BX75', 'CB75', 'CF75', 'CJ75', 'CN75', 'CR75', 'CV75', 'CZ75']
    if customer.get('capital_amount'):
        _fill_digits(ws, capital_cells, customer['capital_amount'])

    # 法人番号（13桁: DX75,EB75,EF75,EJ75,EN75,ER75,EV75,EZ75,FD75,FH75,FL75,FP75,FT75）
    corp_number_cells = [
        'DX75', 'EB75', 'EF75', 'EJ75', 'EN75', 'ER75', 'EV75',
        'EZ75', 'FD75', 'FH75', 'FL75', 'FP75', 'FT75',
    ]
    if customer.get('corporate_number'):
        _fill_cells(ws, corp_number_cells, customer['corporate_number'])

    # ============================================================
    # 項番14: 兼業の有無
    # ============================================================
    if application.get('side_business'):
        ws['AR78'] = str(application['side_business'])
        if int(application['side_business']) == 2:
            ws['CG79'] = 'なし'
        elif application.get('side_business_type'):
            ws['CG79'] = application['side_business_type']

    # ============================================================
    # 項番15: 許可換えの区分
    # ============================================================
    if application.get('permit_transfer_category'):
        ws['AL82'] = str(application['permit_transfer_category'])

    # ============================================================
    # 項番16: 旧許可番号
    # ============================================================
    # 旧許可番号（6桁: CZ88,DD88,DH88,DL88,DP88,DT88）
    old_permit_cells = ['CZ88', 'DD88', 'DH88', 'DL88', 'DP88', 'DT88']
    if application.get('old_permit_number'):
        _fill_digits(ws, old_permit_cells, application['old_permit_number'])
    if application.get('old_permit_year'):
        _fill_digits(ws, ['EL88', 'EP88'], application['old_permit_year'])
    if application.get('old_permit_month'):
        _fill_digits(ws, ['EW88', 'FA88'], application['old_permit_month'])
    if application.get('old_permit_day'):
        _fill_digits(ws, ['FH88', 'FL88'], application['old_permit_day'])

    # ============================================================
    # 申請者・代理人（行11-16）
    # ============================================================
    # 申請者住所（所在地）
    if application.get('applicant_address'):
        ws['DK11'] = application['applicant_address']
    # 申請者氏名（法人名）
    if application.get('applicant_name'):
        ws['DK12'] = application['applicant_name']
    # 代表者氏名（法人の場合）
    if customer.get('corporation_type') and int(customer['corporation_type']) == 1:
        if customer.get('representative'):
            ws['DK13'] = customer['representative']
    # 申請者代理人
    if application.get('proxy_name'):
        ws['DL16'] = application['proxy_name']

    # ============================================================
    # 連絡先（行97-105）
    # ============================================================
    if application.get('contact_organization'):
        ws['B97'] = application['contact_organization']
    if application.get('contact_name'):
        ws['AZ97'] = application['contact_name']
    if application.get('contact_phone'):
        ws['B101'] = application['contact_phone']
    if application.get('contact_fax'):
        ws['B105'] = application['contact_fax']

    # バイトデータとして返す
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
