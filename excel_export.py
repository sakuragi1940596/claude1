import copy
import os
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

# 業種に対応するExcelセル列（行33の列）
BUSINESS_TYPE_COLUMNS = [
    'AR', 'AV', 'AZ', 'BD', 'BH', 'BL', 'BP', 'BT', 'BX', 'CB',
    'CF', 'CJ', 'CN', 'CQ', 'CV', 'CZ', 'DD', 'DH', 'DL', 'DP',
    'DT', 'DX', 'EB', 'EF', 'EJ', 'EN', 'ER', 'EV', 'EZ',
]


def generate_excel(application, customer):
    """申請データからExcelファイルを生成してバイトデータを返す"""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb['様式第一号']

    # 申請日（結合セルの開始セルに書き込む）
    if application['application_date']:
        parts = application['application_date'].split('-')
        if len(parts) == 3:
            year = int(parts[0]) - 2018  # 西暦→令和変換
            ws['EZ9'] = str(year)   # FC9はEZ9:FE9の結合内
            ws['FJ9'] = str(int(parts[1]))  # FN9はFJ9:FP9の結合内
            ws['FU9'] = str(int(parts[2]))  # FY9はFU9:FZ9の結合内

    # 申請年月日（項番03）
    if application['application_date']:
        parts = application['application_date'].split('-')
        if len(parts) == 3:
            year = int(parts[0]) - 2018
            ws['AR28'] = f'令和　{year}'
            ws['BF28'] = str(int(parts[1]))
            ws['BQ28'] = str(int(parts[2]))

    # 許可番号（項番01）
    if application['governor_or_minister'] == 2:
        # 知事許可の場合
        pass
    # 般/特（項番01）
    if application['general_or_specific'] == 1:
        ws['CH21'] = '般'
    else:
        ws['CH21'] = '特'

    # 許可番号
    if application['permit_number']:
        ws['CW21'] = f"第　{application['permit_number']}"

    # 許可年月日
    if application['permit_year']:
        ws['EE21'] = f"令和　{application['permit_year']}"
    if application['permit_month']:
        ws['FD21'] = application['permit_month']
    if application['permit_day']:
        ws['FO21'] = application['permit_day']

    # 申請の区分（項番02）
    if application['application_category']:
        ws['AL25'] = str(application['application_category'])

    # 許可の有効期間の調整
    if application['validity_adjustment']:
        ws['FJ25'] = str(application['validity_adjustment'])

    # 許可を受けようとする建設業（項番04）
    if application['business_types']:
        selected = application['business_types'].split(',')
        for i, (label, code) in enumerate(BUSINESS_TYPES):
            if code in selected:
                col = BUSINESS_TYPE_COLUMNS[i]
                ws[f'{col}34'] = '1'  # 許可を受けようとする業種行

    # 既に許可を受けている建設業（項番05）
    if application['existing_business_types']:
        selected = application['existing_business_types'].split(',')
        for i, (label, code) in enumerate(BUSINESS_TYPES):
            if code in selected:
                col = BUSINESS_TYPE_COLUMNS[i]
                ws[f'{col}37'] = '1'

    # 商号又は名称のフリガナ（項番06）
    if customer['name_kana']:
        ws['AR40'] = customer['name_kana']

    # 商号又は名称（項番07）
    if customer['name']:
        ws['AR46'] = customer['name']

    # 代表者氏名のフリガナ（項番08）
    if customer['representative_kana']:
        ws['AR52'] = customer['representative_kana']

    # 代表者又は個人の氏名（項番09）
    if customer['representative']:
        ws['AR55'] = customer['representative']

    # 主たる営業所の所在地 市区町村コード（項番10）
    if application['city_code']:
        ws['AR57'] = application['city_code']

    # 都道府県名
    if customer['prefecture']:
        ws['BN58'] = customer['prefecture']

    # 市区町村名
    if customer['city']:
        ws['DS58'] = customer['city']

    # 主たる営業所の所在地（項番11）
    if customer['address']:
        ws['AR61'] = customer['address']

    # 郵便番号（項番12）
    if customer['postal_code']:
        parts = customer['postal_code'].split('-')
        if len(parts) == 2:
            ws['AR67'] = parts[0]
            ws['BD67'] = f'－'
            ws['BH67'] = parts[1]

    # 電話番号
    if customer['phone']:
        ws['CE67'] = customer['phone']

    # FAX
    if customer['fax']:
        ws['AR70'] = customer['fax']

    # 法人又は個人の別（項番13）
    if customer['corporation_type']:
        ws['AL75'] = str(customer['corporation_type'])

    # 資本金額
    if customer['capital_amount']:
        ws['BT74'] = customer['capital_amount']

    # 法人番号
    if customer['corporate_number']:
        ws['DX74'] = customer['corporate_number']

    # 兼業の有無（項番14）
    if application['side_business']:
        ws['AL78'] = str(application['side_business'])
    if application['side_business_type']:
        ws['CG77'] = application['side_business_type']

    # 許可換えの区分（項番15）
    if application['permit_transfer_category']:
        ws['AL82'] = str(application['permit_transfer_category'])

    # 申請者
    if application['applicant_name']:
        ws['CX13'] = application['applicant_name']
    if application['applicant_address']:
        ws['DL15'] = application['applicant_address']

    # 申請者代理人
    if application['proxy_name']:
        ws['DL16'] = application['proxy_name']

    # 連絡先
    if application['contact_organization']:
        ws['B97'] = application['contact_organization']
    if application['contact_name']:
        ws['AZ97'] = application['contact_name']
    if application['contact_phone']:
        ws['B101'] = application['contact_phone']
    if application['contact_fax']:
        ws['B105'] = application['contact_fax']

    # バイトデータとして返す
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
