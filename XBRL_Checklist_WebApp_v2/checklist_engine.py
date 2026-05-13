"""
checklist_engine.py  — XML 아웃풋 시트 번호 기준 완전 구현 (29개 체크)

출력 시트 번호 ↔ 로직 대응 (Alteryx XML DbFileOutput 기준):
  1-1  Gross 계정 사용 검토
  1-2  초과적립액(과소적립액) 텍사노미 사용 검토
  1-3  재고자산 세부내역 표 (GrossCarryingAmountMember / AllowanceForCreditLossesMember)
  1-4  유동/비유동 축 검토
  2-1  (만료) 대손충당금 멤버
  2-2  (만료) 금융자산 손상차손 축
  2-3  대출약정 텍사노미 검토
  2-4  미착품 텍사노미 검토
  2-5  배당금 텍사노미 검토            ← 추가 항목
  2-6  평균유효세율 검토 (분반기)
  3-1  Axis & Domain & Member 정합성 검토
  3-2  공시금액의 사용 적정성 검토
  4-1  현금흐름 관련 표 내에서 다른 요소 사용
  4-2  현금흐름 관련 표의 전용요소가 다른 표에서 사용
  4-3  판매관리비 관련 표 내에서 다른 요소 사용
  4-4  판매비와관리비 관련 표의 전용요소가 다른 표에서 사용
  4-5  특수관계자 관련 표 내에서 다른 요소 사용
  4-6  특수관계자 관련 표의 전용요소가 다른 표에서 사용
  5-1  Percent 소숫점 자리수 검토
  5-2  보유하는 주식수 속성 검토
  5-3  이연법인세부채(자산) 텍사노미 및 부호 검토
  5-4  기본주당이익/희석주당이익 속성 검토
  5-5  기초/기말 영문명 검토            ← 추가 항목
  5-6  단위표시 검토                    ← 추가 항목
  6-1  축 확장 검토
  6-2  멤버 합계열 확장 검토
  6-3  Duration / Instant 속성 검토    ← 추가 항목
  7-1  Client Negate 검토
  7-2  현금흐름표 영업활동 현금흐름 검토
"""

import os
import pandas as pd
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Set

# ─── Taxonomy 로드 (모듈 로드 시 1회) (5.5일 추가) ───────────────────────────────
_TAXONOMY_PATH = os.path.join(os.path.dirname(__file__), 'data', 'DART_Negate_Check.xlsx')
try:
    taxonomy = pd.read_excel(_TAXONOMY_PATH)
    taxonomy = taxonomy[['Taxonomy ID', 'DART_Negate']]
    negate = set(taxonomy[taxonomy['DART_Negate'] == 'negate']['Taxonomy ID'])
except Exception:
    taxonomy = pd.DataFrame(columns=['Taxonomy ID', 'DART_Negate'])
    negate = set()


# ─── Axis_Domain_Check 로드 (모듈 로드 시 1회) ───────────────────────────────
_AXIS_CHECK_PATH = os.path.join(os.path.dirname(__file__), 'data', 'Axis_Domain_Check.xlsx')
try:
    axis_check_df = pd.read_excel(_AXIS_CHECK_PATH)[
        ['Table_Number', 'Definition', 'Axis_Domain', 'Axis_Name', 'KEY']
    ].dropna(subset=['KEY'])
    # 검색용 리스트로 미리 변환 (iterrows 반복 제거)
    axis_check_records = axis_check_df.to_dict('records')
    axis_check_keys    = [str(r['KEY']) for r in axis_check_records]
except Exception:
    axis_check_df      = pd.DataFrame(columns=['Table_Number', 'Definition', 'Axis_Domain', 'Axis_Name', 'KEY'])
    axis_check_records = []
    axis_check_keys    = []


# ─── 표 코드 상수 ─────────────────────────────────────────────────────────────
CF_DIRECT_TABLES     = {'D851100', 'D851105'}
CF_INDIRECT_TABLES   = {'D520000', 'D520005'}
CF_TABLES            = {'D851100', 'D851105', 'DX520000', 'DX520005',
                         'D510000', 'D510005', 'DI520000', 'DI520005',
                         'D520000', 'D520005'}
SGA_TABLES           = {'D834310', 'D834315', 'DX830000'}
RELATED_PARTY_TABLES = {'D818000', 'D818005', 'DX837000'}
EPS_TABLES           = {'D838000', 'D838005'}
CAPITAL_TABLES       = {'D861200', 'D861205'}
EQUITY_STMT_TABLES   = {'D610000', 'D610005'}
TAX_TABLES           = {'D835110', 'D835115'}
INVENTORY_NEW_TABLES = {'D826380', 'D826385'}
PENSION_TABLES       = {'D834480', 'D834485'}

# 2-5: 배당금 deprecated 요소 (Node 604 기준)
DIVIDEND_DEPRECATED = {
    'DividendsPaidPreferredSharesPerShare',
    'DividendsPaidOrdinarySharesPerShare',
    'DividendsPaid',
    'DividendsRecognisedAsDistributionsToOwnersPerShare',
    'DividendsPayableOrdinarySharesPerShare',
    'DividendsPayablePreferredSharesPerShare',
    'DividendsProposedOrDeclaredBeforeFinancialStatementsAuthorisedForIssueButNotRecognisedAsDistributionToOwners',
    'DividendsProposedOrDeclaredBeforeFinancialStatementsAuthorisedForIssueButNotRecognisedAsDistributionToOwnersPerShare',
}

# 1-1: Gross 예외 (Node 184 기준)
GROSS_EXCEPTIONS = {'GrossProfit', 'GrossLoanCommitments'}

# 5-4: EPS 요소명 패턴
EPS_NAMES = ('BasicEarningsLossPerShare', 'DilutedEarningsLossPerShare',
             'BasicEarningsPerShare', 'DilutedEarningsPerShare')

# 2-6: 평균유효세율 패턴
EFFECTIVE_TAX_PATTERNS = ['AverageEffectiveTaxRate', 'EffectiveTaxRateFromDiscontinuedOperations']

# 7-2: 영업활동 현금흐름 LineItems 헤더
CF_OPERATING_LINEITEM = 'CashFlowsFromUsedInOperatingActivitiesLineItems'


# ─── 결과 클래스 ──────────────────────────────────────────────────────────────
@dataclass
class CheckIssue:
    role_uri: str = ''; role_code: str = ''
    role_name_ko: str = ''; role_name_en: str = ''
    is_consolidated: Optional[bool] = None
    element_name: str = ''; label_ko: str = ''; label_en: str = ''
    prefix: str = ''; data_type: str = ''; balance: str = ''
    period: str = ''; gubn: str = ''; depth: int = 0
    parent_name: str = ''; parent_label_ko: str = ''; parent_gubn: str = ''
    reason: str = ''
    client_negate: str = ''; dart_negate: str = ''
    table_name_ko: str = ''


@dataclass
class CheckResult:
    check_id: str; title: str; description: str; category: str; sheet: str
    issues: List[CheckIssue] = field(default_factory=list)

    @property
    def issue_count(self): return len(self.issues)
    @property
    def passed(self): return not self.issues
    @property
    def consol_count(self): return sum(1 for i in self.issues if i.is_consolidated is True)
    @property
    def sep_count(self): return sum(1 for i in self.issues if i.is_consolidated is False)


def _mk(row: dict, reason: str, data) -> CheckIssue:
    name = row.get('Name', '')
    el = data.elements.get(name)
    return CheckIssue(
        role_uri=row.get('role_uri', ''), role_code=row.get('role_code', ''),
        role_name_ko=row.get('role_name_ko', ''), role_name_en=row.get('role_name_en', ''),
        is_consolidated=row.get('is_consolidated'),
        element_name=name,
        label_ko=row.get('Label(KO)') or (el.label_ko if el else ''),
        label_en=row.get('Label(EN)') or (el.label_en if el else ''),
        prefix=row.get('Prefix', ''), data_type=row.get('DataType', ''),
        balance=row.get('Balance', ''), period=row.get('Period', ''),
        gubn=row.get('구분', ''), depth=row.get('depth', 0),
        parent_name=row.get('parent', ''), parent_label_ko=row.get('parent_label_ko', ''),
        parent_gubn=row.get('parent_gubn', ''), reason=reason,
        table_name_ko=row.get('table_name_ko', ''),
    )


def run_all_checks(data) -> OrderedDict:
    rows = data.presentation_rows
    return OrderedDict([
        ('1-1', _c1_1(rows, data)),
        ('1-2', _c1_2(rows, data)),
        ('1-3', _c1_3(rows, data)),
        ('1-4', _c1_4(rows, data)),
        ('2-1', _c2_1(rows, data)),
        ('2-2', _c2_2(rows, data)),
        ('2-3', _c2_3(rows, data)),
        ('2-4', _c2_4(rows, data)),
        ('2-5', _c2_5(rows, data)),   # 배당금 (추가)
        ('2-6', _c2_6(rows, data)),   # 평균유효세율
        ('3-1', _c3_1(data.axis_domain_rows, data)),   # Axis & Member 정합성
        ('3-2', _c3_2(rows, data)),   # 공시금액 적정성
        ('4-1', _c4_1(rows, data)),
        ('4-2', _c4_2(rows, data)),
        ('4-3', _c4_3(rows, data)),
        ('4-4', _c4_4(rows, data)),
        ('4-5', _c4_5(rows, data)),
        ('4-6', _c4_6(rows, data)),
        ('5-1', _c5_1(rows, data)),
        ('5-2', _c5_2(rows, data)),
        ('5-3', _c5_3(rows, data)),
        ('5-4', _c5_4(rows, data)),
        ('5-5', _c5_5(rows, data)),   # 기초/기말 (추가)
        ('5-6', _c5_6(rows, data)),   # 단위표시 (추가)
        ('6-1', _c6_1(rows, data)),
        ('6-2', _c6_2(rows, data)),
        ('6-3', _c6_3(rows, data)),   # Duration/Instant (추가)
        ('7-1', _c7_1(rows, data)),   # Negate
        ('7-2', _c7_2(rows, data)),   # CF 영업활동
    ])


# ═════════════════════════════════════════════════════════════════════════════
# 1. 특정요소 사용검토  ← 기존 대비 수정됨 (_c1_2 pandas 기반 재작성)
# ═════════════════════════════════════════════════════════════════════════════

def _c1_1(rows, data):
    """1-1: Gross 계정 사용 검토
    Net(순액) 텍사노미 사용이 원칙이며, Gross를 사용할 시 검토한다.
    예외: GrossProfit, GrossLoanCommitments

    Alteryx 로직 (Node 562→554→184):
      GrossProfit/GrossLoanCommitments 제외
      !IsNull([Gross Account])  : 표준 택소노미 Gross Account 룩업에 걸리는 요소
      Contains([Label Role], "total") : totalLabel 사용 요소 추가 필터
    """
    r = CheckResult('1-1', 'Gross 계정 사용 검토',
        'Net(순액) 텍사노미 사용이 원칙입니다. '
        'Name에 "Gross"가 포함되고 totalLabel인 요소를 검출합니다. '
        '예외: GrossProfit, GrossLoanCommitments.',
        '특정요소 사용검토', 'Checklist_1-1')
    for row in rows:
        name = row.get('Name', '')
        if any(exc in name for exc in GROSS_EXCEPTIONS):
            continue
        if 'Gross' not in name:
            continue
        lbl_role = row.get('Label Role', '').lower()
        if 'total' in lbl_role:
            r.issues.append(_mk(row, 'Gross 계정 totalLabel 사용 — Net 사용 검토 필요', data))
        elif name == 'GrossCarryingAmountMember':
            r.issues.append(_mk(row, 'GrossCarryingAmountMember 사용 — Gross 계정 검토 필요', data))
    return r


def _c1_2(rows, data):
    """1-2: 초과적립액(과소적립액) 텍사노미 사용 검토
    Alteryx 로직 (Node 418→415→563):
      Filter: TABLE_NUMBER in PENSION_TABLES (D834480/D834485)
      flag_1: Name contains "DefinedBenefitObligationAtPresentValue"
      flag_2: Name contains "PlanAssetsAtFairValue"
      flag_3: Name contains "SurplusDeficitInPlan"
      GroupBy table_name_ko → Sum flags
      Filter(418): Sum_flag_1>1 OR Sum_flag_2>1 OR Sum_flag_3>1
      Filter(415): Sum_flag_1+Sum_flag_2+Sum_flag_3 < 6
      Inner Join → Filter FOOTNOTES 제외 → 비확장 item SurplusDeficitInPlan 검출
    """
    r = CheckResult('1-2', '초과적립액(과소적립액) 텍사노미 사용 검토',
        'SurplusDeficitInPlan이 DefinedBenefitObligationAtPresentValue/PlanAssetsAtFairValue와 '
        'Set 형태로 사용되지 않은 경우 검출합니다.',
        '특정요소 사용검토', 'Checklist_1-2')

    # Step 1: 퇴직급여 표 필터
    pension_rows = [r_ for r_ in rows if r_.get('TABLE_NUMBER') in PENSION_TABLES]
    if not pension_rows:
        return r

    df = pd.DataFrame(pension_rows)

    # Step 2: flag 생성
    df['flag_1'] = df['Name'].str.contains('DefinedBenefitObligationAtPresentValue', na=False).astype(int)
    df['flag_2'] = df['Name'].str.contains('PlanAssetsAtFairValue', na=False).astype(int)
    df['flag_3'] = df['Name'].str.contains('SurplusDeficitInPlan', na=False).astype(int)

    # Step 3: table_name_ko 기준 groupby & sum
    grp = df.groupby('table_name_ko')[['flag_1', 'flag_2', 'flag_3']].sum().reset_index()
    grp.columns = ['table_name_ko', 'sum_f1', 'sum_f2', 'sum_f3']

    # Step 4 & 5: Filter(418) AND Filter(415)
    left = grp[
        ((grp['sum_f1'] > 1) | (grp['sum_f2'] > 1) | (grp['sum_f3'] > 1)) &
        ((grp['sum_f1'] + grp['sum_f2'] + grp['sum_f3']) < 6)
    ][['table_name_ko']]

    # Step 6: Inner join (Right = pension_rows 전체, Left = 위 필터 결과)
    joined = left.merge(df, on='table_name_ko', how='inner')

    # Step 7: FOOTNOTES 제외
    joined = joined[joined['구분'] != 'FOOTNOTES']

    # Step 8 & 9: 비확장 item SurplusDeficitInPlan → 검토대상
    filtered = joined[
        (joined['확장여부'] != '확장') &
        (joined['Element'] == 'item') &
        (joined['Name'] == 'SurplusDeficitInPlan')
    ]
    for _, row in filtered.iterrows():
        r.issues.append(_mk(row.to_dict(),
            'SurplusDeficitInPlan 단독 사용 — DefinedBenefitObligation/PlanAssets와 Set 형태 필요', data))

    return r


def _c1_3(rows, data):
    """1-3: 재고자산 세부내역 표 검토
    D826380/D826385 표에서 GrossCarryingAmountMember 옆에
    AllowanceForInventoryValuationMember(평가충당금) 대신
    AllowanceForCreditLossesMember(손실충당금)가 사용된 경우를 검출.

    Alteryx 로직 (Node 365→397→399):
      Contains([Role Definition], "[D826380]") OR Contains([Role Definition], "[D826385]")
      [Name] = "GrossCarryingAmountMember"
      PrevValue/NextValue = "AllowanceForCreditLossesMember"
    """
    r = CheckResult('1-3', '재고자산 세부내역 표 검토',
        '재고자산 표(D826380/D826385)에서 GrossCarryingAmountMember 인접 요소를 검토합니다. '
        '평가충당금(AllowanceForInventoryValuationMember) 사용이 원칙이며, '
        '손실충당금(AllowanceForCreditLossesMember) 사용 시 검출합니다.',
        '특정요소 사용검토', 'Checklist_1-3')

    inv_rows = [rw for rw in rows if rw.get('TABLE_NUMBER') in INVENTORY_NEW_TABLES]

    # GrossCarryingAmountMember 가 있는 (role_uri, TABLE_NUMBER) 집합
    gross_keys: Set[tuple] = set()
    for rw in inv_rows:
        if rw.get('Name') == 'GrossCarryingAmountMember':
            gross_keys.add((rw.get('role_uri', ''), rw.get('TABLE_NUMBER', '')))

    if not gross_keys:
        return r

    # AllowanceForCreditLossesMember 가 같은 표에 있으면 검출
    for rw in inv_rows:
        key = (rw.get('role_uri', ''), rw.get('TABLE_NUMBER', ''))
        if key in gross_keys and rw.get('Name') == 'AllowanceForCreditLossesMember':
            r.issues.append(_mk(rw,
                'GrossCarryingAmountMember 표에서 AllowanceForCreditLossesMember(손실충당금) 사용 — '
                'AllowanceForInventoryValuationMember(평가충당금) 사용 필요', data))
    return r


def _c1_4(rows, data):
    """1-4: 유동/비유동 축 검토
    Alteryx 로직 (Node 419):
      contains([Prefix], "entity") AND Contains([Name], "Axis") AND
      (Contains([Label(KO)], "유동") OR Contains([Label(KO)], "비유동"))
    """
    r = CheckResult('1-4', '유동/비유동 축 검토',
        'entity prefix의 Axis 요소 중 Label(KO)에 "유동" 또는 "비유동"이 포함된 경우 검출합니다.',
        '특정요소 사용검토', 'Checklist_1-4')
    for row in rows:
        lbl = row.get('Label(KO)', '')
        if (row.get('Prefix', '').startswith('entity')
                and 'Axis' in row.get('Name', '')
                and ('유동' in lbl or '비유동' in lbl)):
            r.issues.append(_mk(row, '유동/비유동 확장 Axis 사용 — 표준 축 사용 검토', data))
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 2. 텍사노미 검토
# ═════════════════════════════════════════════════════════════════════════════

def _c2_1(rows, data):
    """2-1: (만료 텍사노미) 대손충당금 멤버 사용 검토
    Alteryx 로직: contains([Name], "AllowanceForCreditLossesMember")
    """
    r = CheckResult('2-1', '(만료 텍사노미) 대손충당금 멤버 사용 검토',
        'Name에 AllowanceForCreditLossesMember가 포함된 만료 요소를 검출합니다.',
        '텍사노미 검토', 'Checklist_2-1')
    for row in rows:
        if 'AllowanceForCreditLossesMember' in row.get('Name', ''):
            r.issues.append(_mk(row, 'AllowanceForCreditLossesMember 만료 요소 사용', data))
    return r


def _c2_2(rows, data):
    """2-2: (만료 텍사노미) 금융자산의 손상차손 축 사용 검토
    Alteryx 로직: contains([Name], "ImpairmentOfFinancialAssetsAxis")
    """
    r = CheckResult('2-2', '(만료 텍사노미) 금융자산의 손상차손 축 사용 검토',
        'Name에 ImpairmentOfFinancialAssetsAxis가 포함된 만료 Axis를 검출합니다.',
        '텍사노미 검토', 'Checklist_2-2')
    for row in rows:
        if 'ImpairmentOfFinancialAssetsAxis' in row.get('Name', ''):
            r.issues.append(_mk(row, 'ImpairmentOfFinancialAssetsAxis 만료 Axis 사용', data))
    return r


def _c2_3(rows, data):
    """2-3: 대출약정 텍사노미 검토
    Alteryx 로직 (Node 388): Contains([Name], "LoanCommitments")
    """
    r = CheckResult('2-3', '대출약정 텍사노미 검토',
        'Name에 "LoanCommitments"가 포함된 요소를 검출합니다.',
        '텍사노미 검토', 'Checklist_2-3')
    for row in rows:
        if 'LoanCommitments' in row.get('Name', ''):
            r.issues.append(_mk(row, '대출약정 만료 텍사노미 요소 사용', data))
    return r


def _c2_4(rows, data):
    """2-4: 미착품 텍사노미 검토
    Alteryx 로직 (Node 170→171→192):
      Contains([Role Definition], "[D826380]") OR Contains([Role Definition], "[D826385]")
      Contains([Label(KO)], "미착")
      !Contains([Label(EN)], "CurrentInventoriesInTransit")
    """
    r = CheckResult('2-4', '미착품 텍사노미 검토',
        '재고자산 표(D826380/D826385)에서 "미착" 항목의 Label(EN)이 '
        'CurrentInventoriesInTransit이 아닌 경우 검출합니다.',
        '텍사노미 검토', 'Checklist_2-4')
    for row in rows:
        if row.get('TABLE_NUMBER') not in INVENTORY_NEW_TABLES:
            continue
        if '미착' in row.get('Label(KO)', ''):
            if 'CurrentInventoriesInTransit' not in row.get('Label(EN)', ''):
                r.issues.append(_mk(row, '미착품 → CurrentInventoriesInTransit 사용 필요', data))
    return r


def _c2_5(rows, data):
    """2-5: 배당금 텍사노미 검토  ← 추가 항목
    Alteryx 로직 (Node 604):
      [Name] IN (DIVIDEND_DEPRECATED 목록)
    """
    r = CheckResult('2-5', '배당금 텍사노미 검토',
        '만료 배당금 요소 사용 시 검출합니다. '
        'DividendsPaid를 포함한 8가지 만료 배당금 텍사노미를 대상으로 합니다.',
        '텍사노미 검토', 'Checklist_2-5')
    for row in rows:
        if row.get('Name', '') in DIVIDEND_DEPRECATED:
            r.issues.append(_mk(row, '만료 배당금 텍사노미 요소 사용', data))
    return r


def _c2_6(rows, data):
    """2-6: 평균유효세율 검토 (분반기)
    Alteryx 로직 (Node 392): Contains([Name], "AverageEffectiveTaxRate")
    법인세 표(D835110/D835115) 외부에서 사용된 경우 검출.
    """
    r = CheckResult('2-6', '평균유효세율 검토 (분반기)',
        'AverageEffectiveTaxRate 요소가 법인세 표(D835110/D835115) 외부에서 사용된 경우 검출합니다.',
        '텍사노미 검토', 'Checklist_2-6')
    for row in rows:
        if (any(p in row.get('Name', '') for p in EFFECTIVE_TAX_PATTERNS)
                and row.get('TABLE_NUMBER') not in TAX_TABLES):
            r.issues.append(_mk(row, '평균유효세율 요소가 법인세 표 외부에서 사용', data))
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 3. 축-멤버 정합성 검토  ← 기존 대비 수정됨 (_c3_1, _c3_2 전면 재작성)
# ═════════════════════════════════════════════════════════════════════════════

def _c3_1(rows, data):
    """3-1: Axis & Domain & Member 정합성 검토
    Alteryx 로직 (FindReplace):
      Find Within Field: axis_domain_rows[KEY_axis]
      Find Value:        Axis_Domain_Check[KEY]  (Any Part of Field)
      → KEY 매칭 결과를 KEY2로 붙여 불일치 검출
    """
    r = CheckResult('3-1', 'Axis & Domain & Member 정합성 검토',
        'axis_domain_rows의 KEY 값이 Axis_Domain_Check의 KEY와 일치하지 않는 경우 검출합니다.',
        '축-멤버 정합성 검토', 'Checklist_3-1')

    for row in rows:
        key_val = str(row.get('KEY_axis') or '')

        # FindReplace: Axis_Domain_Check의 KEY가 KEY_axis 안에 포함되는지 확인
        matched = None
        for ref_key, ref in zip(axis_check_keys, axis_check_records):
            if ref_key and ref_key in key_val:
                matched = ref
                break

        key2 = str(matched['KEY']) if matched is not None else None

        # CHECK 판정
        if key2 is None or key_val != str(key2):
            status = 'CHECK'
        else:
            status = 'OK'

        # 최종 검출: CHECK이면서 비확장 멤버
        if (status == 'CHECK'
                and row.get('확장여부') != '확장'
                and row.get('축_도메인') == '멤버'):
            r.issues.append(_mk(row,
                f'KEY: {key_val} / KEY2: {key2} — 축-멤버 구조 검토 필요', data))

    return r


def _c3_2(rows, data):
    """3-2: 공시금액의 사용 적정성 검토
    Alteryx 로직:
      Left:  축_도메인 == "축" → TABLE_NUMBER + table_name_ko groupby → count >= 2
      Right: 구분 not in ("FOOTNOTES", "TABLE")
      Inner Join → Name == "ReportedAmountMember" → 검토대상
    """
    r = CheckResult('3-2', '공시금액의 사용 적정성 검토',
        '축이 2개 이상인 표에서 ReportedAmountMember(공시금액)가 사용된 경우 검출합니다.',
        '축-멤버 정합성 검토', 'Checklist_3-2')

    df = pd.DataFrame(rows)

    # Left: 축_도메인 == "축" 행을 TABLE_NUMBER + table_name_ko로 groupby, count >= 2
    axis_df = df[df['축_도메인'] == '축']
    group_counts = (axis_df
                    .groupby(['role_uri', 'TABLE_NUMBER', 'table_name_ko'])
                    .size()
                    .reset_index(name='axis_count'))
    left = group_counts[group_counts['axis_count'] >= 2][['role_uri', 'TABLE_NUMBER', 'table_name_ko']]

    # Right: 구분 not in ("FOOTNOTES", "TABLE")
    right = df[~df['구분'].isin(['FOOTNOTES', 'TABLE'])]

    # Inner join
    joined = left.merge(right, on=['role_uri', 'TABLE_NUMBER', 'table_name_ko'], how='inner')

    # 검토대상: Name == "ReportedAmountMember"
    filtered = joined[joined['Name'] == 'ReportedAmountMember']
    for _, row in filtered.iterrows():
        r.issues.append(_mk(row.to_dict(),
            '축이 2개 이상인 표에서 ReportedAmountMember(공시금액) 사용 — 구조 검토 필요', data))

    return r


# ═════════════════════════════════════════════════════════════════════════════
# 4. 전용요소 사용 검토
# ═════════════════════════════════════════════════════════════════════════════

def _entity_excl(rows, table_set: set) -> set:
    """특정 표 그룹의 entity(확장) LINEITEM 전용 요소 집합 반환."""
    return {rw['Name'] for rw in rows
            if rw.get('TABLE_NUMBER') in table_set
            and rw.get('구분') == 'LINEITEM'
            and rw.get('Prefix', '').startswith('entity')}


def _c4_1(rows, data):
    """4-1: 현금흐름 관련 표 내에서 다른 요소 사용
    Alteryx: 표준 CF 주석 룩업 → Null(IsEmpty [Name2]) AND item AND 비확장
    판관비/특수관계자 표의 entity 전용요소가 CF 표에서 사용된 경우.
    """
    r = CheckResult('4-1', '현금흐름 관련 표 내에서 다른 요소 사용',
        '판관비/특수관계자 표의 확장 전용 요소가 현금흐름 표에서 사용된 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-1')
    other_excl = _entity_excl(rows, SGA_TABLES) | _entity_excl(rows, RELATED_PARTY_TABLES)
    for row in rows:
        if (row.get('TABLE_NUMBER') in CF_TABLES
                and row.get('Name') in other_excl):
            r.issues.append(_mk(row, '다른 표 전용 요소가 현금흐름 표에서 사용', data))
    return r


def _c4_2(rows, data):
    """4-2: 현금흐름 관련 표의 전용요소가 다른 표에서 사용
    Alteryx: !IsEmpty([Name2]) AND item AND 비확장 (CF 전용 요소가 타 표에)
    """
    r = CheckResult('4-2', '현금흐름 관련 표의 전용요소가 다른 표에서 사용',
        'CF 표의 확장 전용 요소가 CF 표 외부에서 사용된 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-2')
    cf_excl = _entity_excl(rows, CF_TABLES)
    for row in rows:
        if (row.get('TABLE_NUMBER') not in CF_TABLES
                and row.get('Name') in cf_excl):
            r.issues.append(_mk(row, 'CF 전용 요소를 다른 표에서 사용', data))
    return r


def _c4_3(rows, data):
    """4-3: 판매관리비 관련 표 내에서 다른 요소 사용
    Alteryx (Node 435): Contains([Role Definition], "D83431") → SGA 표
    CF/특수관계자 표의 entity 전용요소가 SGA 표에서 사용된 경우.
    """
    r = CheckResult('4-3', '판매관리비 관련 표 내에서 다른 요소 사용',
        'CF/특수관계자 표의 확장 전용 요소가 판관비 표에서 사용된 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-3')
    other_excl = _entity_excl(rows, CF_TABLES) | _entity_excl(rows, RELATED_PARTY_TABLES)
    for row in rows:
        if (row.get('TABLE_NUMBER') in SGA_TABLES
                and row.get('Name') in other_excl):
            r.issues.append(_mk(row, '다른 표 전용 요소가 판관비 표에서 사용', data))
    return r


def _c4_4(rows, data):
    """4-4: 판매비와관리비 관련 표의 전용요소가 다른 표에서 사용
    Alteryx (Node 449): !Contains([Role Definition], "D83431") → SGA 외부
    """
    r = CheckResult('4-4', '판매비와관리비 관련 표의 전용요소가 다른 표에서 사용',
        'SGA 표의 확장 전용 요소가 SGA 표 외부에서 사용된 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-4')
    sga_excl = _entity_excl(rows, SGA_TABLES)
    for row in rows:
        if (row.get('TABLE_NUMBER') not in SGA_TABLES
                and row.get('Name') in sga_excl):
            r.issues.append(_mk(row, 'SGA 전용 요소를 다른 표에서 사용', data))
    return r


def _c4_5(rows, data):
    """4-5: 특수관계자 관련 표 내에서 다른 요소 사용
    Alteryx (Node 453): Contains([Role Definition], "D81800") AND [구분] = "LINEITEM"
    특수관계자 표 LINEITEM 중 ifrs-full prefix + NameOf로 시작하는 요소.
    """
    r = CheckResult('4-5', '특수관계자 관련 표 내에서 다른 요소 사용',
        '특수관계자 표(D818000/D818005) 내 LINEITEM 중 ifrs-full prefix이고 '
        'NameOf로 시작하는 요소를 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-5')
    for row in rows:
        if (row.get('TABLE_NUMBER') in RELATED_PARTY_TABLES
                and row.get('구분') == 'LINEITEM'
                and row.get('Prefix') == 'ifrs-full'
                and row.get('Name', '').startswith('NameOf')):
            r.issues.append(_mk(row, '특수관계자 표 내 NameOf... 요소 사용 — 검토 필요', data))
    return r


def _c4_6(rows, data):
    """4-6: 특수관계자 관련 표의 전용요소가 다른 표에서 사용
    Alteryx (Node 457): !Contains([Role Definition], "D81800") → RP 외부
    """
    r = CheckResult('4-6', '특수관계자 관련 표의 전용요소가 다른 표에서 사용',
        'RP 표의 확장 전용 요소가 RP 표 외부에서 사용된 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-6')
    rp_excl = _entity_excl(rows, RELATED_PARTY_TABLES)
    for row in rows:
        if (row.get('TABLE_NUMBER') not in RELATED_PARTY_TABLES
                and row.get('Name') in rp_excl):
            r.issues.append(_mk(row, 'RP 전용 요소를 다른 표에서 사용', data))
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 5. 속성/데이터타입 검토
# ═════════════════════════════════════════════════════════════════════════════

def _c5_1(rows, data):
    """5-1: Percent 소숫점 자리수 검토
    Alteryx 로직 (Node 377/483): Contains([DataType], "percent")
    """
    r = CheckResult('5-1', 'Percent 소숫점 자리수 검토',
        'DataType에 "percent"가 포함된 모든 요소를 검출합니다. '
        '이자율/할인율 요소의 Decimal 소숫점 자리수 속성 적정성을 검토합니다.',
        '속성/데이터타입 검토', 'Checklist_5-1')
    for row in rows:
        if 'percent' in row.get('DataType', '').lower():
            r.issues.append(_mk(row, 'percentItemType 요소 — 소숫점 자리수(Decimal) 속성 검토 필요', data))
    return r


def _c5_2(rows, data):
    """5-2: 보유하는 주식수 속성 검토
    sharesItemType + EPS표(D838000/D838005) / 자본금표(D861200/D861205).
    단, IncreaseDecreaseInNumberOfSharesOutstanding 제외.
    Alteryx 로직 (Node 483): Contains([DataType], "shares")
    """
    r = CheckResult('5-2', '보유하는 주식수 속성 검토',
        'sharesItemType 요소의 Period/DataType 속성 적정성을 검토합니다. '
        'EPS 표(D838000/D838005)와 자본금 표(D861200/D861205)가 대상입니다.',
        '속성/데이터타입 검토', 'Checklist_5-2')
    SHARE_TABLES = EPS_TABLES | CAPITAL_TABLES
    seen_outstanding: Set[str] = set()
    for row in rows:
        if 'shares' not in row.get('DataType', '').lower():
            continue
        tn = row.get('TABLE_NUMBER', '')
        if tn not in SHARE_TABLES:
            continue
        name = row.get('Name', '')
        if 'IncreaseDecrease' in name:
            continue
        if tn in CAPITAL_TABLES and 'Outstanding' in name:
            key = f'{name}|{tn}'
            if key in seen_outstanding:
                continue
            seen_outstanding.add(key)
        r.issues.append(_mk(row, '주식수 요소 속성 검토 필요 (Period/DataType)', data))
    return r


def _c5_3(rows, data):
    """5-3: 이연법인세부채(자산) 텍사노미 및 부호 검토
    Alteryx 로직 (Node 375/483): [Name] = "DeferredTaxLiabilityAsset"
    """
    r = CheckResult('5-3', '이연법인세부채(자산) 텍사노미 및 부호 검토',
        'DeferredTaxLiabilityAsset 요소 사용 시 검출합니다. '
        'DeferredTaxLiability/DeferredTaxAsset 분리 사용을 권장합니다.',
        '속성/데이터타입 검토', 'Checklist_5-3')
    for row in rows:
        if row.get('Name') == 'DeferredTaxLiabilityAsset':
            r.issues.append(_mk(row, 'DeferredTaxLiabilityAsset 복합 요소 사용 — 부호 검토 필요', data))
    return r


def _c5_4(rows, data):
    """5-4: 기본주당이익/희석주당이익 속성 검토
    Alteryx 로직 (Node 385/483):
      Contains([Name], "BasicEarningsLossPerShare") OR Contains([Name], "DilutedEarningsLossPerShare")
    Period=INSTANT인 경우 검출 (DURATION이어야 함).
    """
    r = CheckResult('5-4', '기본주당이익/희석주당이익 속성 검토',
        'BasicEarningsLossPerShare 또는 DilutedEarningsLossPerShare 요소의 '
        'Period가 INSTANT인 경우 검출합니다. DURATION이어야 합니다.',
        '속성/데이터타입 검토', 'Checklist_5-4')
    for row in rows:
        if (any(ep in row.get('Name', '') for ep in EPS_NAMES)
                and row.get('Period', '').upper() == 'INSTANT'):
            r.issues.append(_mk(row, '주당이익 Period=INSTANT — DURATION 필요', data))
    return r


def _c5_5(rows, data):
    """5-5: 기초/기말 영문명 검토  ← 추가 항목
    Alteryx 로직 (Node 597):
      (Contains([Label(KO)], "기초") OR Contains([Label(KO)], "기말")
       OR Contains([Label(EN)], "Begin") OR Contains([Label(EN)], "Ending"))
      AND [구분] = "LINEITEM"
    """
    r = CheckResult('5-5', '기초/기말 영문명 검토',
        'Label(KO)에 "기초"/"기말"이 포함되거나 '
        'Label(EN)에 "Begin"/"Ending"이 포함된 LINEITEM 요소를 검출합니다. '
        'Opening/Closing 영문명 사용 여부를 검토합니다.',
        '속성/데이터타입 검토', 'Checklist_5-5')
    for row in rows:
        if row.get('구분') != 'LINEITEM':
            continue
        lko = row.get('Label(KO)', '')
        len_ = row.get('Label(EN)', '')
        if ('기초' in lko or '기말' in lko
                or 'Begin' in len_ or 'Ending' in len_):
            r.issues.append(_mk(row, '기초/기말 영문명 검토 필요 (Opening/Closing 사용 권장)', data))
    return r


def _c5_6(rows, data):
    """5-6: 단위표시 검토  ← 추가 항목
    각 표(Role/TABLE_NUMBER)별로 monetaryItemType 요소가 0건인 경우 '단위미표시' 검출.

    Alteryx 로직 (Node 619→620→621):
      단위표시구분 = IF DataType = "monetaryItemType" THEN "단위표시숫자" ELSE "단위표시 불필요"
      Summarize by (Role Definition, TABLE NAME): SUM(CNT)
      IF CNT_단위표시숫자 = 0 THEN "단위미표시"
    """
    r = CheckResult('5-6', '단위표시 검토',
        '각 표(Role/TABLE NUMBER)별로 monetaryItemType 요소가 없는 경우(단위미표시)를 검출합니다. '
        'LINEITEM이 있으나 금액형 요소가 0건인 표를 확인합니다.',
        '속성/데이터타입 검토', 'Checklist_5-6')

    # 표별 집계
    table_monetary: Dict[tuple, int] = defaultdict(int)
    table_lineitem:  Dict[tuple, int] = defaultdict(int)
    table_meta:      Dict[tuple, dict] = {}
    SKIP_GUBN = {'TABLE', 'FOOTNOTES', 'Axis', 'Domain', 'Member'}

    for row in rows:
        if row.get('구분') in SKIP_GUBN:
            continue
        key = (row.get('role_uri', ''), row.get('TABLE_NUMBER', ''),
               row.get('role_code', ''), row.get('role_name_ko', ''),
               row.get('is_consolidated'))
        table_lineitem[key] += 1
        if 'monetary' in row.get('DataType', '').lower():
            table_monetary[key] += 1
        if key not in table_meta:
            table_meta[key] = row

    for key, cnt_li in table_lineitem.items():
        if cnt_li == 0:
            continue
        if table_monetary[key] == 0:
            base = table_meta[key]
            iss = CheckIssue(
                role_uri=key[0], role_code=key[2],
                role_name_ko=key[3], role_name_en=base.get('role_name_en', ''),
                is_consolidated=key[4],
                element_name=key[1] or key[2],
                label_ko=f'LINEITEM {cnt_li}건 중 monetaryItemType: 0건',
                label_en='', prefix='', data_type='', balance='', period='',
                gubn='TABLE', depth=0, parent_name='', parent_label_ko='', parent_gubn='',
                reason='단위미표시 — 해당 표에 monetaryItemType 요소가 없음',
            )
            r.issues.append(iss)
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 6. 확장 검토  ← 기존 대비 수정됨 (_c6_1, _c6_2 필터 조건 변경)
# ═════════════════════════════════════════════════════════════════════════════

def _c6_1(rows, data):
    """6-1: 축 확장 검토 — 축은 확장하지 않는다
    Alteryx 로직 (Node 368→369):
      contains([Prefix], "entity")
      Contains([Name], "Axis")
    """
    r = CheckResult('6-1', '축 확장 검토',
        'Name에 "Axis"가 포함되며 entity prefix인 확장 Axis 요소를 검출합니다. '
        '축(Axis)은 확장하지 않는 것이 원칙입니다.',
        '확장 검토', 'Checklist_6-1')
    for row in rows:
        if ('entity' in row.get('Prefix', '')
                and 'Axis' not in row.get('Name', '')
                and row.get('구분', '') == 'Axis'):
            r.issues.append(_mk(row, '확장(entity) Axis 요소 — Name에 Axis 미포함', data))
    return r


def _c6_2(rows, data):
    """6-2: 멤버 합계열 확장 검토 — 합계열은 확장하지 않는다
    Alteryx 로직 (Node 174→175):
      Contains([Label(KO)], "합계")
      Contains([Prefix], "entity") AND [Element] = "Member"
    """
    r = CheckResult('6-2', '멤버 합계열 확장 검토',
        'Label(KO)에 "합계"가 포함되고 entity prefix인 확장 Member 요소를 검출합니다. '
        '합계열이 필요한 경우 도메인의 합계열(Yes)을 사용해야 합니다.',
        '확장 검토', 'Checklist_6-2')
    for row in rows:
        if ('합계' in row.get('Label(KO)', '')
                and row.get('Prefix', '').startswith('entity')
                and row.get('Element', '') == 'Member'):
            r.issues.append(_mk(row, '합계열 멤버 확장 사용 — 도메인 합계열 사용 필요', data))
    return r


def _c6_3(rows, data):
    """6-3: Duration / Instant 속성 검토  ← 추가 항목
    Alteryx 로직 (Node 599):
      Contains([Prefix], "entity") AND [구분] = "LINEITEM"
      AND Left([Name], 4) != "Title"
    확장 LINEITEM 요소의 Period(Duration/Instant) 속성 적정성 검토.
    """
    r = CheckResult('6-3', 'Duration / Instant 속성 검토',
        'entity(확장) prefix의 LINEITEM 요소 전체를 검출합니다. '
        'Period 속성(Duration/Instant)이 표준 요소와 일치하는지 검토합니다. '
        '"Title"로 시작하는 요소는 제외합니다.',
        '확장 검토', 'Checklist_6-3')
    for row in rows:
        if (row.get('Prefix', '').startswith('entity')
                and row.get('구분') == 'LINEITEM'
                and not row.get('Name', '').startswith('Title')):
            r.issues.append(_mk(row, '확장(entity) LINEITEM — Duration/Instant 속성 검토 필요', data))
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 7. 기타  ← 기존 대비 수정됨 (_c7_1 pandas 기반 재작성, _c7_2 필터 조건 변경)
# ═════════════════════════════════════════════════════════════════════════════

def _c7_1(rows, data):
    """7-1: Client Negate 검토
    Alteryx 로직 (Node 587):
      Client_Negate != DART_Negate → 검토대상
    DART Taxonomy의 preferredLabel 기준 negated 여부와 입력 파일의 Label Role 비교.
    """
    r = CheckResult('7-1', 'Client Negate 검토',
        '하기 추출 리스트들은 기본 속성이 Negated가 아닌 내역을 사용자가 negated로 설정한 내역들을 확인할 수 있습니다.',
        '기타', 'Checklist_7-1')

    if not rows:
        return r

    # ── 1. Client DataFrame 클렌징 ──────────────────────────────────────────
    client = pd.DataFrame(rows)
    client['Taxonomy ID'] = client['Prefix'].fillna('') + '_' + client['Name'].fillna('')
    client['Client_Negate'] = client['Label Role'].apply(
        lambda x: 'negate' if 'negated' in str(x).lower() else '-'
    )

    # ── 2. DART negate DataFrame 준비 ───────────────────────────────────────
    negate_df = taxonomy[taxonomy['DART_Negate'] == 'negate'][['Taxonomy ID', 'DART_Negate']].drop_duplicates()

    # ── 3. Client 기준 Left join → Negate_Check ─────────────────────────────
    Negate_Check = client.merge(negate_df, on='Taxonomy ID', how='left').reset_index(drop=True)
    Negate_Check['DART_Negate'] = Negate_Check['DART_Negate'].fillna('-')

    # ── 4. 검토여부 ─────────────────────────────────────────────────────────
    Negate_Check['검토여부'] = Negate_Check.apply(
        lambda x: '비검토대상' if x['Client_Negate'] == x['DART_Negate'] else '검토대상', axis=1
    )

    # ── 5. 정렬 ─────────────────────────────────────────────────────────────
    # 구조내려받기 엑셀 시트 순서 그대로 사용
    sheet_order_map = {s: i for i, s in enumerate(dict.fromkeys(client['Sheet']))}
    Negate_Check['_sort_sheet'] = Negate_Check['Sheet'].map(sheet_order_map).fillna(9999)
    # 연결(0) → 별도(1)
    Negate_Check['_sort_consol'] = Negate_Check['is_consolidated'].map(
        {True: 0, False: 1}
    ).fillna(2)
    Negate_Check = Negate_Check.sort_values(
        ['_sort_consol', '_sort_sheet']
    ).reset_index(drop=True)

    # ── 7. 검토대상만 필터 → CheckIssue 변환 (원래 열만 노출) ─────────────────
    filtered = Negate_Check[
        (Negate_Check['Client_Negate'] == 'negate') &
        (Negate_Check['검토여부'] == '검토대상')
    ]
    for _, row in filtered.iterrows():
        iss = _mk(row.to_dict(), 'Client_Negate ≠ DART_Negate — Negate 적용 여부 검토 필요', data)
        iss.client_negate = row['Client_Negate']
        iss.dart_negate   = row['DART_Negate']
        r.issues.append(iss)

    return r


def _c7_2(rows, data):
    """7-2: 현금흐름표 영업활동 현금흐름 검토
    Alteryx 로직 (Node 423→428→424→429):
      Filter(423): TABLE_NUMBER가 D851100 또는 D851105인 라인
      Multi-Row Formula(428): NextValue = 다음 행의 Name
      Filter(424): Name에 CashFlowsFromUsedInOperatingActivitiesLineItems 미포함 라인
      Filter(429): NextValue가 ProfitLoss가 아닌 라인 → 검토대상
    """
    r = CheckResult('7-2', '현금흐름표 영업활동 현금흐름 검토',
        '현금흐름표 직접법(D851100/D851105)에서 '
        'CashFlowsFromUsedInOperatingActivitiesLineItems를 제외한 라인 중 '
        'NextValue가 ProfitLoss가 아닌 경우를 검출합니다.',
        '기타', 'Checklist_7-2')

    # Filter(423): TABLE_NUMBER가 D851100 또는 D851105('CF_DIRECT_TABLES')인 라인
    table_groups: Dict[tuple, List[dict]] = defaultdict(list)
    for row in rows:
        if row.get('TABLE_NUMBER') in CF_DIRECT_TABLES:  # {'D851100', 'D851105'}
            key = (row.get('role_uri', ''), row.get('TABLE_NUMBER', ''))
            table_groups[key].append(row)

    for key, grp_rows in table_groups.items():
        for i, row in enumerate(grp_rows):
            # Multi-Row Formula(428): NextValue = 다음 행의 Name
            next_name = grp_rows[i + 1].get('Name', '') if i + 1 < len(grp_rows) else ''

            # Filter(424): Name에 CashFlowsFromUsedInOperatingActivitiesLineItems 포함된 라인만
            if CF_OPERATING_LINEITEM not in row.get('Name', ''):
                continue

            # Filter(429): NextValue가 ProfitLoss가 아닌 라인만
            if next_name == 'ProfitLoss':
                continue

            r.issues.append(_mk(row,
                f'NextValue: {next_name} — 영업활동 현금흐름 구조 검토 필요', data))
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 요약
# ═════════════════════════════════════════════════════════════════════════════

SECTION_LABELS = {
    '1': '특정요소 사용검토',
    '2': '텍사노미 검토',
    '3': '축-멤버 정합성 검토',
    '4': '전용요소 사용 검토',
    '5': '속성/데이터타입 검토',
    '6': '확장 검토',
    '7': '기타',
}


def get_summary(results: OrderedDict) -> dict:
    cats: Dict[str, dict] = {}
    for cid, res in results.items():
        sec = cid.split('-')[0]
        cat = SECTION_LABELS.get(sec, '기타')
        if cat not in cats:
            cats[cat] = {'checks': [], 'total_issues': 0, 'section': int(sec)}
        cats[cat]['checks'].append(res)
        cats[cat]['total_issues'] += res.issue_count
    cats = dict(sorted(cats.items(), key=lambda x: x[1].get('section', 99)))
    return {
        'total_checks':       len(results),
        'total_issues':       sum(r.issue_count for r in results.values()),
        'checks_with_issues': sum(1 for r in results.values() if not r.passed),
        'categories':         cats,
    }
