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


# ─── 0511 데이터 파일 로드 (모듈 로드 시 1회) ────────────────────────────────
_DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')

_TAXONOMY_PATH = os.path.join(_DATA_DIR, 'DART_Negate_Check.xlsx')
try:
    taxonomy = pd.read_excel(_TAXONOMY_PATH)
    taxonomy = taxonomy[['Taxonomy ID', 'DART_Negate']]
    negate = set(taxonomy[taxonomy['DART_Negate'] == 'negate']['Taxonomy ID'])
except Exception:
    taxonomy = pd.DataFrame(columns=['Taxonomy ID', 'DART_Negate'])
    negate = set()

_AXIS_CHECK_PATH = os.path.join(_DATA_DIR, 'Axis_Domain_Check.xlsx')
try:
    axis_check_df = pd.read_excel(_AXIS_CHECK_PATH)[
        ['Table_Number', 'Definition', 'Axis_Domain', 'Axis_Name', 'KEY']
    ].dropna(subset=['KEY'])
    axis_check_records = axis_check_df.to_dict('records')
    axis_check_keys    = [str(r['KEY']) for r in axis_check_records]
except Exception:
    axis_check_df      = pd.DataFrame(columns=['Table_Number', 'Definition', 'Axis_Domain', 'Axis_Name', 'KEY'])
    axis_check_records = []
    axis_check_keys    = []


# ─── 표 코드 상수 ─────────────────────────────────────────────────────────────
CF_DIRECT_TABLES     = {'D851100', 'D851105'}
CF_INDIRECT_TABLES   = {'D520000', 'D520005'}
EPS_TABLES           = {'D838000', 'D838005'}
CAPITAL_TABLES       = {'D861200', 'D861205'}
EQUITY_STMT_TABLES   = {'D610000', 'D610005'}
PENSION_TABLES       = {'D834480', 'D834485'}
INVENTORY_NEW_TABLES = {'D826380','D826385'}

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

# 1-1: Gross Account 목록 (Find-Replace R 인풋)
GROSS_ACCOUNT_LIST = {
    'IntangibleAssetsUnderDevelopmentGross',
    'BuildingsGross',
    'BuildingIncidentalFacilitiesGross',
    'ConstructionInProgressGross',
    'ToolsAndEquipmentGross',
    'SupplyFacilitiesGross',
    'ReceivablesOnConstructionContracts',
    'ExchangeableBonds',
    'StructureGross',
    'FinanceLeaseAssetGross',
    'MachineryGross',
    'OtherIntangibleAssetsGross',
    'OtherPropertyPlantAndEquipmentGross',
    'OtherInventoriesGross',
    'CurrentFinanceLeaseReceivablesGross',
    'ShortTermLoans',
    'ShortTermTradeReceivable',
    'ShortTermOtherReceivables',
    'ShortTermAccruedIncome',
    'ShortTermDueFromCustomersForContractWork',
    'ShortTermDepositsProvided',
    'ShortTermPrepaidConstructionCosts',
    'ShortTermAdvancePayments',
    'ShortTermPrepaidExpenses',
    'ShortTermDeferredAncillaryIncomeForLoans',
    'LeaseholdDeposits',
    'ReceivablesAgent',
    'LicencesAndFranchisesGross',
    'RentalAssetGross',
    'IntangibleExplorationAndEvaluationAssetsGross',
    'UnfinishedProgramGross',
    'GoodsInTransitGross',
    'ByProductGross',
    'ReceivablesRealestateSales',
    'BrandNamesGross',
    'NonCurrentEmissionRightGross',
    'NonCurrentBiologicalAssetsGross',
    'OfficeEquipmentGross',
    'UsufructContributionAssetGross',
    'MerchandiseGross',
    'ShipsGross',
    'ProductionSuppliesGross',
    'FittingGross',
    'ExperimentMaterialGross',
    'BondWithWarrant',
    'UtilityModelRightsGross',
    'GoodwillGross',
    'FinishedHousingGross',
    'FinishedProgramGross',
    'LotGross',
    'RawMaterialsGross',
    'CurrentEmissionRight',
    'CurrentBiologicalAssetsGross',
    'TangibleExplorationAndEvaluationAssetsGross',
    'RentStructureGross',
    'GrossCapitalisedResearchAndDevelopmentExpenseForBioindustry',
    'LongTermReceivablesOnConstructionContracts',
    'LongTermAdvancesOnConstructionContracts',
    'LongTermContractReserve',
    'NoncurrentFinanceLeaseReceivablesGross',
    'LongTermOtherGuaranteeDepositReceivedGross',
    'LongTermLoansGross',
    'LongTermReceivablesAgent',
    'LongTermTradePayablesGross',
    'LongTermTradeAndOtherNonCurrentReceivablesGross',
    'LongTermTradeReceivablesGross',
    'LongTermOtherReceivablesGross',
    'LongTermAccruedIncomeGross',
    'LongTermOtherPayablesGross',
    'LongTermAccruedExpensesGross',
    'LongTermDueFromCustomersForContractWork',
    'LongTermDepositsProvidedGross',
    'LongTermReceivablesRealestateSales',
    'LongTermAdvancePaymentsGross',
    'LongTermPrepaidExpenses',
    'LongTermAdvancesCustomers',
    'LongTermRentReceivedInAdvance',
    'LongTermWithholdingsBanks',
    'LongTermGuaranteeDepositWithholdings',
    'LongTermDeferredAncillaryIncomeForLoans',
    'DeferredIncomeClassifiedAsNoncurrent',
    'LongTermGuaranteeDepositRentGross',
    'LongTermLeaseholdDeposits',
    'LongTermBorrowingsGross',
    'WorkInProgressGross',
    'CopyrightsPatentsAndOtherIndustrialPropertyRightsServiceAndOperatingRightsGross',
    'SuppliesGross',
    'ElectronicAutomationDevelopmentGross',
    'ElectronicFacilitiesGross',
    'TelexAndTelephoneSubscriptionRights',
    'ConvertibleBonds',
    'ConvertibleRedeemablePreferredStockLiabilities',
    'FinishedGoodsGross',
    'MastheadsAndPublishingTitlesGross',
    'RecipesFormulaeModelsDesignsAndPrototypesGross',
    'FixturesAndFittingsGross',
    'VehiclesGross',
    'LandRightGross',
    'MiningRightsGross',
    'ComputerSoftwareGross',
    'LandGross',
    'LandUseRightGross',
    'InvestmentPropertyGross',
    'CashAndCashEquivalentsGross',
}

# 5-4: EPS 요소명 패턴
EPS_NAMES = ('BasicEarningsLossPerShare', 'DilutedEarningsLossPerShare',
             'BasicEarningsPerShare', 'DilutedEarningsPerShare')

# 7-2: 영업활동 현금흐름 LineItems 헤더
CF_OPERATING_LINEITEM = 'CashFlowsFromUsedInOperatingActivitiesLineItems'


# ─── 결과 클래스 ──────────────────────────────────────────────────────────────
@dataclass
class CheckIssue:
    role_uri: str = ''; role_code: str = ''
    role_name_ko: str = ''; role_name_en: str = ''
    is_consolidated: Optional[bool] = None
    element_name: str = ''; label_ko: str = ''; label_en: str = ''
    label_role: str = ''
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
        label_role=row.get('Label Role', ''),
        prefix=row.get('Prefix', ''), data_type=row.get('DataType', ''),
        balance=row.get('Balance', ''), period=row.get('Period', ''),
        gubn=row.get('구분', ''), depth=row.get('depth', 0),
        parent_name=row.get('parent', ''), parent_label_ko=row.get('parent_label_ko', ''),
        parent_gubn=row.get('parent_gubn', ''), reason=reason,
        client_negate=row.get('Client_Negate', ''),
        table_name_ko=row.get('table_name_ko', ''),
    )


def _clean(row_dict: dict) -> dict:
    """pandas iterrows() row의 NaN / numpy 타입을 _mk() 호환 값으로 변환."""
    import math
    result = {}
    for k, v in row_dict.items():
        if isinstance(v, float) and math.isnan(v):
            result[k] = None if k == 'is_consolidated' else ''
        elif hasattr(v, 'item'):   # numpy scalar → Python native
            result[k] = v.item()
        else:
            result[k] = v
    return result


def run_all_checks(data, std=None) -> OrderedDict:
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
        ('3-1', _c3_1(data.axis_domain_rows, data)),  # Axis & Member 정합성 (0511: axis_domain_rows)
        ('3-2', _c3_2(rows, data)),   # 공시금액 적정성
        ('4-1', _c4_1(rows, data, std)),
        ('4-2', _c4_2(rows, data, std)),
        ('4-3', _c4_3(rows, data, std)),
        ('4-4', _c4_4(rows, data, std)),
        ('4-5', _c4_5(rows, data, std)),
        ('4-6', _c4_6(rows, data, std)),
        ('5-1', _c5_1(rows, data)),
        ('5-2', _c5_2(rows, data)),
        ('5-3', _c5_3(rows, data)),
        ('5-4', _c5_4(rows, data)),
        ('5-5', _c5_5(rows, data)),   # 기초/기말 (추가)
        ('5-6', _c5_6(rows, data)),   # 단위표시 (추가)
        ('6-1', _c6_1(rows, data)),
        ('6-2', _c6_2(rows, data)),
        ('6-3', _c6_3(rows, data)),   # Duration/Instant (추가)
        ('7-1', _c7_1(rows, data)),   # Negate (0511: pandas 기반)
        ('7-2', _c7_2(rows, data)),   # CF 영업활동
    ])


# ═════════════════════════════════════════════════════════════════════════════
# 1. 특정요소 사용검토
# ═════════════════════════════════════════════════════════════════════════════

def _c1_1(rows, data):
    """1-1: Gross 계정 사용 검토
    Net(순액) 텍사노미 사용이 원칙이며, Gross를 사용할 시 검토한다.
    예외: GrossProfit, GrossLoanCommitments

    Alteryx 로직:
      1. Contains([Name], "GrossProfit") OR Contains([Name], "GrossLoanCommitments") → F (예외 제외)
      2. Find-Replace: Name을 GROSS_ACCOUNT_LIST 로 룩업 → !IsNull([Gross Account]) → T
      3. Contains([Label Role], "total") → F → 이슈
    """
    r = CheckResult('1-1', 'Gross 계정 사용 검토',
        'Net(순액) 텍사노미 사용이 원칙입니다. '
        'Gross Account 목록에 있는 요소 중 totalLabel인 경우 검출합니다. '
        '예외: GrossProfit, GrossLoanCommitments.',
        '특정요소 사용검토', 'Checklist_1-1')
    for row in rows:
        name = row.get('Name', '')
        if any(exc in name for exc in GROSS_EXCEPTIONS):
            continue
        if name not in GROSS_ACCOUNT_LIST:
            continue
        if 'total' not in row.get('Label Role', '').lower():
            r.issues.append(_mk(row, 'Gross 계정 totalLabel 미사용 — Net 사용 검토 필요', data))
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

    pension_rows = [r_ for r_ in rows if r_.get('TABLE_NUMBER') in PENSION_TABLES]
    if not pension_rows:
        return r

    df = pd.DataFrame(pension_rows)

    df['flag_1'] = df['Name'].str.contains('DefinedBenefitObligationAtPresentValue', na=False).astype(int)
    df['flag_2'] = df['Name'].str.contains('PlanAssetsAtFairValue', na=False).astype(int)
    df['flag_3'] = df['Name'].str.contains('SurplusDeficitInPlan', na=False).astype(int)

    grp = df.groupby('table_name_ko')[['flag_1', 'flag_2', 'flag_3']].sum().reset_index()
    grp.columns = ['table_name_ko', 'sum_f1', 'sum_f2', 'sum_f3']

    left = grp[
        ((grp['sum_f1'] > 1) | (grp['sum_f2'] > 1) | (grp['sum_f3'] > 1)) &
        ((grp['sum_f1'] + grp['sum_f2'] + grp['sum_f3']) < 6)
    ][['table_name_ko']]

    joined = left.merge(df, on='table_name_ko', how='inner')
    joined = joined[joined['구분'] != 'FOOTNOTES']

    filtered = joined[
        (joined['확장여부'] != '확장') &
        (joined['Element'] == 'item') &
        (joined['Name'] == 'SurplusDeficitInPlan')
    ]
    for _, row in filtered.iterrows():
        r.issues.append(_mk(_clean(row.to_dict()),
            'SurplusDeficitInPlan 단독 사용 — DefinedBenefitObligation/PlanAssets와 Set 형태 필요', data))

    return r


def _c1_3(rows, data):
    """1-3: 재고자산 세부내역 표 검토
    Alteryx 로직 (Node 365→395/396→367→397→399):
      1. TABLE_NUMBER in {D826380, D826385} 필터
      2. PrevValue = 이전 행 Name (Multi-Row Formula 395)
      3. NextValue = 다음 행 Name (Multi-Row Formula 396)
      4. Name == "GrossCarryingAmountMember" 필터 (Filter 367)
      5. NOT(
           (PrevValue=="CarryingAmountMember" AND NextValue=="AllowanceForInventoryValuationMember")
           OR
           (NextValue=="CarryingAmountMember" AND PrevValue=="AllowanceForInventoryValuationMember")
         ) → False 추출 (Filter 397)
      6. NextValue == "AllowanceForCreditLossesMember" 추출 (Filter 399)
    """
    r = CheckResult('1-3', '재고자산 세부내역 표 검토',
        '재고자산 표(D826380/D826385)에서 GrossCarryingAmountMember 전후에 '
        'CarryingAmountMember + AllowanceForInventoryValuationMember 구조가 없고 '
        'AllowanceForCreditLossesMember(손실충당금)가 인접한 경우를 검출합니다.',
        '특정요소 사용검토', 'Checklist_1-3')

    inv_rows = [rw for rw in rows if rw.get('TABLE_NUMBER') in INVENTORY_NEW_TABLES]
    if not inv_rows:
        return r

    df = pd.DataFrame(inv_rows)

    all_filtered = []
    for _, grp in df.groupby('role_uri', sort=False):
        grp = grp.reset_index(drop=True)

        # Step 2 & 3: PrevValue / NextValue 생성
        grp['PrevValue'] = grp['Name'].shift(1).fillna('')
        grp['NextValue'] = grp['Name'].shift(-1).fillna('')

        # Step 4: GrossCarryingAmountMember 행만
        gross = grp[grp['Name'] == 'GrossCarryingAmountMember'].copy()
        if gross.empty:
            continue

        # Step 5: 정상 구조(CarryingAmountMember ↔ AllowanceForInventoryValuationMember 인접) 제외
        cond_ok = (
            ((gross['PrevValue'] == 'CarryingAmountMember') &
             (gross['NextValue'] == 'AllowanceForInventoryValuationMember')) |
            ((gross['NextValue'] == 'CarryingAmountMember') &
             (gross['PrevValue'] == 'AllowanceForInventoryValuationMember'))
        )
        not_ok = gross[~cond_ok]

        # Step 6: NextValue == AllowanceForCreditLossesMember
        filtered = not_ok[not_ok['NextValue'] == 'AllowanceForCreditLossesMember']
        all_filtered.append(filtered)

    if all_filtered:
        result_df = pd.concat(all_filtered, ignore_index=True)
        for _, row in result_df.iterrows():
            r.issues.append(_mk(_clean(row.to_dict()),
                'GrossCarryingAmountMember 인접 행이 AllowanceForCreditLossesMember(손실충당금) — '
                'AllowanceForInventoryValuationMember(평가충당금) 사용 필요', data))
    return r


def _c1_4(rows, data):
    """1-4: 유동/비유동 축 검토
    Alteryx 로직 (Node 419→595→Filter):
      1. Prefix에 "entity" 포함 AND Name에 "Axis" 포함
         AND Label(KO)에 "유동" 또는 "비유동" 포함 (Filter 419)
      2. Element == "Axis" 인 라인만 추출 (Filter 595)
      3. Name에 "ClassesOfCurrentAndNoncurrentValuesUsedForChangesInAssetAndLiabilityAxis"
         포함되지 않는 라인만 추출
    """
    r = CheckResult('1-4', '유동/비유동 축 검토',
        'entity prefix의 Axis 요소 중 Label(KO)에 "유동"/"비유동"이 포함되고, '
        'ClassesOfCurrentAndNoncurrentValuesUsedForChangesInAssetAndLiabilityAxis가 아닌 경우 검출합니다.',
        '특정요소 사용검토', 'Checklist_1-4')

    EXCLUDE_NAME = 'ClassesOfCurrentAndNoncurrentValuesUsedForChangesInAssetAndLiabilityAxis'

    for row in rows:
        lbl  = row.get('Label(KO)', '')
        name = row.get('Name', '')

        # Step 1
        if not ('entity' in row.get('Prefix', '')
                and 'Axis' in name
                and ('유동' in lbl or '비유동' in lbl)):
            continue
        # Step 2
        if row.get('Element', '') != 'Axis':
            continue
        # Step 3
        if EXCLUDE_NAME in name:
            continue

        r.issues.append(_mk(row, '유동/비유동 확장 Axis 사용 — 표준 축 사용 검토 필요', data))
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
    Alteryx 로직:
      1. Contains([Role Definition], "[D827580]") OR Contains([Role Definition], "[D827585]")
      2. Contains([Label(KO)], "대출약정")
      3. Contains([Name], "LoanCommitments")
      4. [확장여부] = "확장"
    """
    r = CheckResult('2-3', '대출약정 텍사노미 검토',
        '대출약정 표(D827580/D827585)에서 Label(KO)에 "대출약정"이 포함되고 '
        'Name에 "LoanCommitments"가 포함되지 않은 확장 요소를 검출합니다.',
        '텍사노미 검토', 'Checklist_2-3')
    for row in rows:
        table_num = row.get('TABLE_NUMBER', '')
        # 1. D827580/D827585 표만 대상
        if not ('D827580' in table_num or 'D827585' in table_num):
            continue
        # 2. '대출약정' 포함 이름
        labelko = row.get('Label(KO)', '')
        if '대출약정' not in labelko:
            continue
        # 3. 'LoanCommitments' 제외 이름
        name = row.get('Name', '')
        if 'LoanCommitments' in name:
            continue
        if row.get('확장여부') == '확장':
            r.issues.append(_mk(row, '대출약정 표에서 LoanCommitments 확장 요소 사용', data))
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
        table_num = row.get('TABLE_NUMBER', '')
        # 1. D826380/D826385 표만 대상
        if not ('D826380' in table_num or 'D826385' in table_num):
            continue
        # 2. '미착' 포함 이름
        labelko = row.get('Label(KO)', '')
        if '미착' not in labelko:
            continue
        # 3. 'CurrentInventoriesInTransit' 포함 이름
        labelen = row.get('Label(EN)', '')
        if 'CurrentInventoriesInTransit' in labelko:
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
    Alteryx 로직 (T-T-F):
      1. Contains([Role Definition], "[D835110]") OR Contains([Role Definition], "[D835115]")  → T
      2. Contains([Label(KO)], "유효세율")  → T
      3. Contains([Name], "AverageEffectiveTaxRate")  → F (불일치 = 이슈)
    법인세 표 내에서 Label(KO)에 "유효세율"이 포함되나 Name이 AverageEffectiveTaxRate가 아닌 경우 검출.
    """
    r = CheckResult('2-6', '평균유효세율 검토 (분반기)',
        '법인세 표(D835110/D835115)에서 Label(KO)에 "유효세율"이 포함되나 '
        'Name이 AverageEffectiveTaxRate가 아닌 경우 검출합니다.',
        '텍사노미 검토', 'Checklist_2-6')
    for row in rows:
        table_num = row.get('TABLE_NUMBER', '')        
        # 1. D835110/D835115 표만 대상
        if not ('D835110' in table_num or 'D835115' in table_num):
            continue
        # 2. '유효세율' 포함 이름
        labelko = row.get('Label(KO)', '')
        if '유효세율' not in labelko:
            continue
        # 3. 'AverageEffectiveTaxRate' 제외 이름
        name = row.get('Name', '')
        if 'AverageEffectiveTaxRate' not in name:
            r.issues.append(_mk(row, '법인세 표 내 유효세율 라벨이지만 AverageEffectiveTaxRate 미사용', data))
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 3. 축-멤버 정합성 검토
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

        matched = None
        for ref_key, ref in zip(axis_check_keys, axis_check_records):
            if ref_key and ref_key in key_val:
                matched = ref
                break

        key2 = str(matched['KEY']) if matched is not None else None

        if key2 is None or key_val != str(key2):
            status = 'CHECK'
        else:
            status = 'OK'

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

    axis_df = df[df['축_도메인'] == '축']
    group_counts = (axis_df
                    .groupby(['role_uri', 'TABLE_NUMBER', 'table_name_ko'])
                    .size()
                    .reset_index(name='axis_count'))
    left = group_counts[group_counts['axis_count'] >= 2][['role_uri', 'TABLE_NUMBER', 'table_name_ko']]

    right = df[~df['구분'].isin(['FOOTNOTES', 'TABLE'])]
    joined = left.merge(right, on=['role_uri', 'TABLE_NUMBER', 'table_name_ko'], how='inner')

    filtered = joined[joined['Name'] == 'ReportedAmountMember']
    for _, row in filtered.iterrows():
        r.issues.append(_mk(_clean(row.to_dict()),
            '축이 2개 이상인 표에서 ReportedAmountMember(공시금액) 사용 — 구조 검토 필요', data))

    return r


# ═════════════════════════════════════════════════════════════════════════════
# 4. 전용요소 사용 검토  (표준 택사노미 룩업 기반)
# ═════════════════════════════════════════════════════════════════════════════

def _is_std_item(row: dict) -> bool:
    """비확장(non-entity) item 요소인지 확인."""
    return (row.get('Element', '') == 'item'
            and row.get('확장여부') != '확장')


def _c4_1(rows, data, std=None):
    """4-1: 현금흐름 관련 표 내에서 다른 요소 사용
    Alteryx 로직:
      1. (Contains([Role Definition], "D85110") or Contains([Role Definition], "D52000"))
         AND [구분] = "LINEITEM" → T
      2. Contains([Name], "Adjustments") → F (Adjustments 제외, 없는 것만 유지)
      3. Find-Replace Source: Axis_Domain_Check.xlsx에서
         [Table_Number] IN ("D851100","D851105","DX520000","DX520005",
                            "D510000","D510005","DI520000","DI520005","D520000","D520005")
         Name 기준 매칭 → Name2 채움 (= std.cf_excl)
      4. IsEmpty([Name2]) AND [Element]='item' AND [확장여부]!='확장' → T → 이슈
    """
    r = CheckResult('4-1', '현금흐름 관련 표 내에서 다른 요소 사용',
        'CF 표(D851100/D520000) 내 비확장 item 요소 중 '
        'Axis_Domain_Check 기준 CF 전용 요소가 아닌 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-1')
    if not std or not std.cf_excl:
        return r
    for row in rows:
        # 1. D85110/D52000 계열 표의 LINEITEM 행만 대상
        table_num = row.get('TABLE_NUMBER', '')
        if not ('D85110' in table_num or 'D52000' in table_num):
            continue
        if row.get('구분') != 'LINEITEM':
            continue
        # 2. Adjustments 포함 이름 제외 (F 분기 유지)
        name = row.get('Name', '')
        if 'Adjustments' in name:
            continue
        # 3+4. IsEmpty([Name2]): Name이 std.cf_excl에 없으면 → item이고 비확장이면 이슈
        if name not in std.cf_excl:
            if _is_std_item(row):
                r.issues.append(_mk(row, 'CF 표에서 CF 전용 요소가 아닌 요소 사용', data))
    return r


def _c4_2(rows, data, std=None):
    """4-2: 현금흐름 관련 표의 전용요소가 다른 표에서 사용
    Alteryx 로직:
      1. Contains([Role Definition], "D851100") OR Contains([Role Definition], "D520000") → F (CF 표 제외)
      2. Find-Replace: element_in_cf → T → 이슈
    """
    r = CheckResult('4-2', '현금흐름 관련 표의 전용요소가 다른 표에서 사용',
        'CF 표(D851100/D520000) 외부에서 표준 텍사노미 기준 CF 전용 요소를 사용한 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-2')
    if not std or not std.cf_excl:
        return r
    for row in rows:
        # 1. D85110/D52000 계열 제외 표 LINEITEM 행만 대상
        table_num = row.get('TABLE_NUMBER', '')
        if ('D85110' in table_num or 'D52000' in table_num):
            continue
        if row.get('구분') != 'LINEITEM':
            continue
        # 2. Adjustments 포함 이름
        name = row.get('Name', '')
        if 'Adjustments' not in name:
            continue
        # 3+4. !IsEmpty([Name2]): Name이 std.cf_excl에 있으면 → item이고 비확장이면 이슈
        if name in std.cf_excl:
            if _is_std_item(row):
                r.issues.append(_mk(row, 'CF 전용 표준 요소를 CF 표 외부에서 사용', data))
    return r


def _c4_3(rows, data, std=None):
    """4-3: 판매관리비 관련 표 내에서 다른 요소 사용
    Alteryx 로직:
      1. Contains([Role Definition], "D834310") → T
      2. Contains([Name], "Adjustments") → F (제외)
      3. Find-Replace: Name을 표준 SGA 텍사노미 요소로 룩업
      4. IsEmpty([Name2]) AND [Element]='item' AND [확장여부]!='확장' → T → 이슈
    """
    r = CheckResult('4-3', '판매관리비 관련 표 내에서 다른 요소 사용',
        '판관비 표(D834310) 내 비확장 item 요소 중 '
        '표준 텍사노미 SGA 전용 요소가 아닌 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-3')
    if not std or not std.sga_excl:
        return r
    for row in rows:
        # 1. D83431 표의 LINEITEM 행만 대상
        if 'D83431' not in row.get('TABLE_NUMBER', ''):
            continue
        if row.get('구분') != 'LINEITEM':
            continue
        # 2+3. SellingGeneralAdministrativeExpenses / SellingGeneralAndAdministrativeExpenses 포함 이름 제외
        name = row.get('Name', '')
        if 'SellingGeneralAdministrativeExpenses' in name or 'RelatedParties' in name:
            continue
        # 4+5. IsEmpty([Name2]): Name이 std.sga_excl에 없으면 → item이고 비확장이면 이슈
        if name not in std.sga_excl:
            if _is_std_item(row):
                r.issues.append(_mk(row, 'SGA 표에서 표준 SGA 전용 요소가 아닌 요소 사용', data))
    return r


def _c4_4(rows, data, std=None):
    """4-4: 판매비와관리비 관련 표의 전용요소가 다른 표에서 사용
    Alteryx 로직:
      1. Contains([Role Definition], "D834310") → F (SGA 표 제외)
      2. Find-Replace: element_in_sga → T → 이슈
    """
    r = CheckResult('4-4', '판매비와관리비 관련 표의 전용요소가 다른 표에서 사용',
        'SGA 표(D834310) 외부에서 표준 텍사노미 기준 SGA 전용 요소를 사용한 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-4')
    if not std or not std.sga_excl:
        return r
    for row in rows:
        # 1. D83431 제외 표의 LINEITEM 행만 대상
        if 'D83431' in row.get('TABLE_NUMBER', ''):
            continue
        if row.get('구분') != 'LINEITEM':
            continue
        # 2+3. SellingGeneralAdministrativeExpenses / SellingGeneralAndAdministrativeExpenses 포함 이름 
        name = row.get('Name', '')
        if 'SellingGeneralAdministrativeExpenses' not in name or 'RelatedParties' not in name:
            continue
        # 4+5. IsEmpty([Name2]): Name이 std.sga_excl에 있으면 → item이고 비확장이면 이슈
        if name in std.sga_excl:
            if _is_std_item(row):
                r.issues.append(_mk(row, 'SGA 전용 표준 요소를 SGA 표 외부에서 사용', data))
    return r


def _c4_5(rows, data, std=None):
    """4-5: 특수관계자 관련 표 내에서 다른 요소 사용
    Alteryx 로직:
      1. Contains([Role Definition], "D81800") AND [구분] = "LINEITEM" → T
      2. !Contains([Name], "RelatedParty") → T
      3. !Contains([Name], "RelatedParties") → T
      4. Find-Replace Source: Axis_Domain_Check.xlsx에서
         [Table_Number] IN ("D818000","D818005","DX837000") 행의 Name (= std.rp_excl)
         Name 기준 매칭 → Name2 채움
      5. IsEmpty([Name2]) AND [Element]='item' AND [확장여부]!='확장' → T → 이슈
    """
    r = CheckResult('4-5', '특수관계자 관련 표 내에서 다른 요소 사용',
        '특수관계자 표(D81800) 내 비확장 item 요소 중 '
        'RP 전용 표(D818000/D818005/DX837000) 기준 RP 전용 요소가 아닌 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-5')
    if not std or not std.rp_excl:
        return r
    for row in rows:
        # 1. D81800 표의 LINEITEM 행만 대상
        if 'D81800' not in row.get('TABLE_NUMBER', ''):
            continue
        if row.get('구분') != 'LINEITEM':
            continue
        # 2+3. RelatedParty / RelatedParties 포함 이름 제외
        name = row.get('Name', '')
        if 'RelatedParty' in name or 'RelatedParties' in name:
            continue
        # 4+5. IsEmpty([Name2]): Name이 std.rp_excl에 없으면 → item이고 비확장이면 이슈
        if name not in std.rp_excl:
            if _is_std_item(row):
                r.issues.append(_mk(row, 'RP 표에서 RP 전용 요소가 아닌 요소 사용', data))
    return r


def _c4_6(rows, data, std=None):
    """4-6: 특수관계자 관련 표의 전용요소가 다른 표에서 사용
    Alteryx 로직:
      1. Contains([Role Definition], "D818000") → F (RP 표 제외)
      2. Find-Replace: element_in_rp → T → 이슈
    """
    r = CheckResult('4-6', '특수관계자 관련 표의 전용요소가 다른 표에서 사용',
        'RP 표(D818000) 외부에서 표준 텍사노미 기준 특수관계자 전용 요소를 사용한 경우 검출합니다.',
        '전용요소 사용 검토', 'Checklist_4-6')
    if not std or not std.rp_excl:
        return r
    for row in rows:
        # 1. D81800 표 제외 LINEITEM 행만 대상
        if 'D81800' in row.get('TABLE_NUMBER', ''):
            continue
        if row.get('구분') != 'LINEITEM':
            continue
        # 2+3. RelatedParty / RelatedParties 포함 이름 
        name = row.get('Name', '')
        if 'RelatedParty' not in name or 'RelatedParties' not in name:
            continue
        # 4+5. !IsEmpty([Name2]): Name이 std.rp_excl에 있으면 → item이고 비확장이면 이슈
        if name in std.rp_excl:
            if _is_std_item(row):
                r.issues.append(_mk(row, 'RP 전용 표준 요소를 RP 표 외부에서 사용', data))
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
    Alteryx 로직:
      1. [Label(KO)]에 "주식수" 또는 "주식 수"를 포함하는 라인 추출 (Filter(379))
      2. [Period]가 "Instant"인 라인 추출 (Filter(381))
      3. 추출된 라인은 검토대상
    """
    r = CheckResult('5-2', '보유하는 주식수 속성 검토',
        'Label(KO)에 "주식수" 또는 "주식 수"가 포함되고 Period가 Instant인 요소를 검출합니다.',
        '속성/데이터타입 검토', 'Checklist_5-2')
    for row in rows:
        labelko = row.get('Label(KO)', '')
        if '주식수' not in labelko and '주식 수' not in labelko:
            continue
        if 'INSTANT' not in row.get('Period', ''):
            r.issues.append(_mk(row, '주식수 요소의 Period=Instant 속성 검토 필요', data))
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
    Alteryx 로직:
      1. [Role Definition]이 [D431410] 또는 [D431415]인 라인 추출 (Filter(382))
      2. [Name]에 "BasicEarningsLossPerShare" 또는 "DilutedEarningsLossPerShare"를 포함하는 라인 추출 (Filter(385))
      3. [Decimal]이 비어있지 않은 라인 추출 (Filter(384))
      4. 추출된 라인은 검토대상
    """
    r = CheckResult('5-4', '기본주당이익/희석주당이익 속성 검토',
        '[D431410]/[D431415] 표에서 BasicEarningsLossPerShare 또는 DilutedEarningsLossPerShare 요소 중 '
        'Decimal이 입력된 요소를 검출합니다.',
        '속성/데이터타입 검토', 'Checklist_5-4')
    EPS_CHECK_NAMES = ('BasicEarningsLossPerShare', 'DilutedEarningsLossPerShare')
    for row in rows:
        table_num = row.get('TABLE_NUMBER', '')
        if 'D431410' not in table_num and 'D431415' not in table_num:
            continue
        name = row.get('Name', '')
        if not any(ep in name for ep in EPS_CHECK_NAMES):
            continue
        if row.get('Decimal', ''):
            r.issues.append(_mk(row, '기본/희석주당이익 요소의 Decimal 속성 검토 필요', data))
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
        labelko  = row.get('Label(KO)', '')
        labelen_l = row.get('Label(EN)', '').lower()
        if not ('기초' in labelko or '기말' in labelko or 'begin' in labelen_l or 'ending' in labelen_l):
            continue
        if row.get('구분') == 'LINEITEM':
            r.issues.append(_mk(row, '기초/기말 영문명 검토 필요 (Opening/Closing 사용 권장)', data))
    return r


def _c5_6(rows, data):
    """5-6: 단위미사용 검토
    Alteryx 로직:
      1. [단위표시구분] = "monetaryItemType"이면 "단위표시숫자", 이외 "단위표시 불필요"
      2. [CNT_단위표시숫자] = if 단위표시구분="단위표시숫자" then 1 else 0
      3. [CNT_단위표시 불필요] = if 단위표시구분="단위표시 불필요" then 1 else 0
      4. [Role Definition]과 [TABLE NAME](table_name_ko)으로 Group By
         → sum CNT_단위표시숫자, sum CNT_단위표시불필요, count 단위표시구분
      5. [단위미표시] = if CNT_단위표시숫자=0 then "단위미표시" else "단위표시"
      6. 최종: [Role Definition], [TABLE NAME], [단위표시/미표시] 추출
    """
    r = CheckResult('5-6', '단위미사용 검토',
        '각 표(Role Definition + TABLE NAME)별로 monetaryItemType 요소가 없는 경우(단위미표시)를 검출합니다.',
        '속성/데이터타입 검토', 'Checklist_5-6')

    if not rows:
        return r

    df = pd.DataFrame(rows)

    # Step 1: 단위표시구분 생성
    df['단위표시구분'] = df['DataType'].apply(
        lambda x: '단위표시숫자' if x == 'monetaryItemType' else '단위표시 불필요'
    )

    # Step 2 & 3: CNT 열 생성
    df['CNT_단위표시숫자']   = (df['단위표시구분'] == '단위표시숫자').astype(int)
    df['CNT_단위표시불필요'] = (df['단위표시구분'] == '단위표시 불필요').astype(int)

    # Step 4: Role Definition + table_name_ko 기준 Group By
    grp_keys = ['Role Definition', 'table_name_ko',
                'role_uri', 'role_code', 'role_name_ko', 'role_name_en', 'is_consolidated']
    grp = (df.groupby(grp_keys, dropna=False)
             .agg(CNT_숫자=('CNT_단위표시숫자', 'sum'),
                  CNT_불필요=('CNT_단위표시불필요', 'sum'))
             .reset_index())

    # Step 5 & 6: CNT_숫자 == 0 → 단위미표시 → TABLE 행 기준으로 _mk() 사용
    미표시 = grp[grp['CNT_숫자'] == 0]
    table_rows_df = df[df['구분'] == 'TABLE']

    for _, row in 미표시.iterrows():
        # 해당 그룹의 TABLE 행을 찾아 _mk() 기준 행으로 사용
        match = table_rows_df[
            (table_rows_df['role_uri'] == row['role_uri']) &
            (table_rows_df['table_name_ko'] == row['table_name_ko'])
        ]
        if not match.empty:
            base_row = _clean(match.iloc[0].to_dict())
        else:
            # TABLE 행이 없으면 그룹 내 첫 번째 행 사용
            first = df[df['role_uri'] == row['role_uri']]
            base_row = _clean(first.iloc[0].to_dict()) if not first.empty else {}

        r.issues.append(_mk(base_row, '단위미표시 — 해당 표에 monetaryItemType 요소 없음', data))
    return r


# ═════════════════════════════════════════════════════════════════════════════
# 6. 확장 검토
# ═════════════════════════════════════════════════════════════════════════════

def _c6_1(rows, data):
    """6-1: 축 확장 검토 — 축은 확장하지 않는다
    Alteryx 로직 (Node 368→369):
      contains([Prefix], "entity")
      Contains([Name], "Axis") — Name에 Axis가 없으면서 구분이 Axis인 경우 검출
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
    """6-3: Duration / Instant 속성 검토
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
# 7. 기타
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

    client = pd.DataFrame(rows)
    client['Taxonomy ID'] = client['Prefix'].fillna('') + '_' + client['Name'].fillna('')
    client['Client_Negate'] = client['Label Role'].apply(
        lambda x: 'negate' if 'negated' in str(x).lower() else '-'
    )

    negate_df = taxonomy[taxonomy['DART_Negate'] == 'negate'][['Taxonomy ID', 'DART_Negate']].drop_duplicates()

    Negate_Check = client.merge(negate_df, on='Taxonomy ID', how='left').reset_index(drop=True)
    Negate_Check['DART_Negate'] = Negate_Check['DART_Negate'].fillna('-')

    Negate_Check['검토여부'] = Negate_Check.apply(
        lambda x: '비검토대상' if x['Client_Negate'] == x['DART_Negate'] else '검토대상', axis=1
    )

    sheet_order_map = {s: i for i, s in enumerate(dict.fromkeys(client['Sheet']))}
    Negate_Check['_sort_sheet'] = Negate_Check['Sheet'].map(sheet_order_map).fillna(9999)
    Negate_Check['_sort_consol'] = Negate_Check['is_consolidated'].map(
        {True: 0, False: 1}
    ).fillna(2)
    Negate_Check = Negate_Check.sort_values(
        ['_sort_consol', '_sort_sheet']
    ).reset_index(drop=True)

    filtered = Negate_Check[
        (Negate_Check['Client_Negate'] == 'negate') &
        (Negate_Check['검토여부'] == '검토대상')
    ]
    for _, row in filtered.iterrows():
        iss = _mk(_clean(row.to_dict()), 'Client_Negate ≠ DART_Negate — Negate 적용 여부 검토 필요', data)
        iss.client_negate = row['Client_Negate']
        iss.dart_negate   = row['DART_Negate']
        r.issues.append(iss)

    return r


def _c7_2(rows, data):
    """7-2: 현금흐름표 영업활동 현금흐름 검토
    Alteryx 로직 (Node 423→428→424→429):
      Filter(423): TABLE_NUMBER가 D851100 또는 D851105('CF_DIRECT_TABLES')인 라인
      Multi-Row Formula(428): NextValue = 다음 행의 Name
      Filter(424): Name에 CashFlowsFromUsedInOperatingActivitiesLineItems 포함된 라인만
      Filter(429): NextValue가 ProfitLoss가 아닌 라인만
    """
    r = CheckResult('7-2', '현금흐름표 영업활동 현금흐름 검토',
        '현금흐름표 직접법(D851100/D851105)에서 '
        'CashFlowsFromUsedInOperatingActivitiesLineItems를 제외한 라인 중 '
        'NextValue가 ProfitLoss가 아닌 경우를 검출합니다.',
        '기타', 'Checklist_7-2')

    table_groups: Dict[tuple, List[dict]] = defaultdict(list)
    for row in rows:
        if row.get('TABLE_NUMBER') in CF_DIRECT_TABLES:
            key = (row.get('role_uri', ''), row.get('TABLE_NUMBER', ''))
            table_groups[key].append(row)

    for key, grp_rows in table_groups.items():
        for i, row in enumerate(grp_rows):
            next_name = grp_rows[i + 1].get('Name', '') if i + 1 < len(grp_rows) else ''

            if CF_OPERATING_LINEITEM not in row.get('Name', ''):
                continue

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
