"""
taxonomy_xlsx_parser.py
IxD 편집기 '구조내려받기' xlsx 파일을 파싱하여
checklist_engine이 사용하는 presentation_rows 형태로 변환
"""

import io, re
from typing import Dict, Optional
import pandas as pd


def _is_consol(text: str) -> Optional[bool]:
    for k in ['Consolidated', 'consolidated', '연결']:
        if k in text: return True
    for k in ['Separated', 'Separate', 'separated', '별도', 'Nonconsolidated']:
        if k in text: return False
    return None


def _extract_table_number(role_def: str) -> str:
    """Role Definition '[D210000] ...' → 'D210000' (가이드 7.비고)"""
    m = re.search(r'\[([A-Za-z]{1,3}X?\d{4,})\]', str(role_def))
    return m.group(1) if m else ''


def _extract_role_code(role_def: str, role_uri: str) -> str:
    code = _extract_table_number(role_def)
    if code: return code
    m = re.search(r'/([A-Z]{1,3}X?\d{4,})$', str(role_uri))
    return m.group(1) if m else ''


def _label_role_short(url: str) -> str:
    s = str(url).strip()
    return '' if not s or s.lower() == 'nan' else s.split('/')[-1]


def _safe(v) -> str:
    if v is None: return ''
    s = str(v).strip()
    return '' if s.lower() == 'nan' else s


def _classify_element(gubn_raw: str, name: str) -> str:
    """가이드 2.2 - Element 분류 (Alteryx 로직 기준)"""
    # Step 1: Name 끝 4글자 기준
    suffix = name[-4:].lower() if len(name) >= 4 else ''
    element = {
        'tory': 'Explanatory', 'ract': 'Abstract',
        'axis': 'Axis',        'lock': 'TextBlock',
        'able': 'Table',       'mber': 'Member',
    }.get(suffix, 'item')
    # Step 2: Name에 'lineitem' 포함 시 override
    if 'lineitem' in name.lower():
        element = 'Lineitem'
    # Step 3: 구분이 FOOTNOTES이면 최종 override
    if gubn_raw.strip().upper() == 'FOOTNOTES':
        element = 'FOOTNOTES'
    return element


def _classify_gubn(gubn_raw: str, name: str) -> str:
    """구분 컬럼 세분화"""
    g = gubn_raw.strip().upper()
    if g == 'TABLE'     or name.endswith('Table'):                   return 'TABLE'
    if g == 'FOOTNOTES' or name.endswith('TextBlock'):               return 'FOOTNOTES'
    if name.endswith('Axis'):                                         return 'Axis'
    if name.endswith('Member'):                                       return 'Member'
    if name.endswith('Domain'):                                       return 'Domain'
    if name.endswith('LineItems') or name.endswith('LineItem'):       return 'LINEITEM'
    if g == 'LINEITEM':                                               return 'LINEITEM'
    if g in ('DOMAIN','MEMBER','AXIS'):                               return g.capitalize()
    return 'LINEITEM'


class TaxonomyXlsxData:
    class _El:
        def __init__(self, lko='', len_='', lr=''):
            self.label_ko=lko; self.label_en=len_; self.label_role=lr; self.abstract=False
    def __init__(self):
        self.company_name=''; self.report_date=''; self.entity_id=''
        self.presentation_rows=[]; self.errors=[]
        self.axis_domain_rows=[]
        self.elements:Dict[str,object]={}
        self.contexts:Dict[str,object]={}
        self.facts=[]; self._fact_elements:set=set()


def parse_taxonomy_xlsx(file_bytes: bytes) -> TaxonomyXlsxData:
    data = TaxonomyXlsxData()
    try:
        xls = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None,
                            header=None, dtype=str, na_filter=False)
    except Exception as e:
        data.errors.append(f'xlsx 읽기 실패: {e}'); return data

    rows = []
    for sheet_name, df in xls.items():
        if '기본정보' in sheet_name:
            _parse_basic_info(df, data); continue

        role_uri = ''; role_def = ''; header_idx = None
        for i, row in df.iterrows():
            v0 = _safe(row.iloc[0]); v1 = _safe(row.iloc[1]) if len(row)>1 else ''
            if 'Role URI' in v0:        role_uri = v1
            elif 'Role Definition' in v0: role_def = v1
            elif v0 == '구분' and i >= 1: header_idx = i; break

        if header_idx is None: continue
        if not role_uri: role_uri = f'sheet:{sheet_name}'

        code      = _extract_role_code(role_def, role_uri)
        table_num = _extract_table_number(role_def) or code
        parts     = role_def.split('|', 1)
        name_ko   = re.sub(r'^\[[^\]]+\]\s*', '', parts[0]).strip()
        name_en   = parts[1].strip() if len(parts)>1 else ''

        # 연결/별도: _is_consol 우선, 없으면 코드 끝자리로 fallback (0=연결, 5=별도)
        is_c = _is_consol(role_def or role_uri)
        if is_c is None and code:
            if code[-1] == '0':   is_c = True
            elif code[-1] == '5': is_c = False

        consol_str = '-'
        if code:
            if code[-1] == '0':   consol_str = '연결'
            elif code[-1] == '5': consol_str = '별도'

        current_table_name_ko = ''
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            def col(j): return _safe(row.iloc[j]) if len(row)>j else ''
            gubn_raw=col(0); prefix=col(1); name=col(2)
            lbl_ko=col(3); lbl_en=col(4); lbl_role_url=col(5)
            dtype=col(6); balance=col(7); period=col(8)
            decimal_val=col(9); fact_val=col(10)

            if not name: continue
            # 컬럼 헤더 행 스킵 (시트 내 반복 헤더)
            if name in ('Name', 'Prefix', '구분', 'Label(KO)', 'Label(EN)',
                        'Label Role', 'DataType', 'Balance', 'Period',
                        'Decimal', 'Fact'): continue

            gubn    = _classify_gubn(gubn_raw, name)
            if gubn == 'TABLE':
                current_table_name_ko = lbl_ko
            element = _classify_element(gubn_raw, name)
            lbl_role= _label_role_short(lbl_role_url)
            period_n= period.upper()     # 가이드: 'INSTANT' / 'DURATION'
            bal_n   = balance.lower()

            # 가이드 기준: 비확장 = '-', 확장 = '확장'
            ext = '확장' if prefix.startswith('entity') else '-'
            client_negate = 'negate' if 'negated' in lbl_role.lower() else '-'
            alias         = '별칭'   if 'terse'   in lbl_role.lower() else '-'
            has_fact      = bool(fact_val) or (decimal_val == '0')

            if name not in data.elements:
                data.elements[name] = TaxonomyXlsxData._El(lbl_ko, lbl_en, lbl_role)

            rows.append({
                'role_uri': role_uri, 'role_code': code,
                'role_name_ko': name_ko, 'role_name_en': name_en,
                'is_consolidated': is_c,
                'Role Definition': role_def,
                'Sheet': sheet_name, '연결/별도': consol_str,
                'Table_Number': table_num, 'TABLE_NUMBER': table_num,
                'depth': 0, 'parent': '', 'parent_label_ko': '', 'parent_gubn': '',
                'Prefix': prefix, 'Name': name,
                'Label(KO)': lbl_ko, 'Label(EN)': lbl_en,
                'Label Role': lbl_role,
                'DataType': dtype, 'Balance': bal_n,
                'Period': period_n,   # 대문자 유지
                'Decimal': decimal_val, 'Fact': fact_val,
                '구분': gubn, 'Element': element,
                '확장여부': ext, 'Client_Negate': client_negate, '별칭여부': alias,
                'PreferredLabel': lbl_role_url,
                'has_fact': has_fact, 'abstract': False,
                'table_name_ko': current_table_name_ko,
            })

    _add_axis_group_fields(rows)
    data.presentation_rows = rows
    data.axis_domain_rows  = [r for r in rows if r.get('GroupID') is not None]
    return data


def _add_axis_group_fields(rows: list):
    """3-1, 3-2 체크용 축-도메인 그룹핑 필드 추가 (role_uri 기준 순차 처리)"""
    from collections import defaultdict

    groups = defaultdict(list)
    for i, row in enumerate(rows):
        groups[row.get('role_uri', '')].append((i, row))

    for _, indexed_rows in groups.items():
        prev_element     = None
        prev_axis_domain = None
        prev_group_id    = None
        prev_axis_name   = None

        for _, (orig_idx, row) in enumerate(indexed_rows):
            element = row.get('Element', '')

            if prev_element == 'Axis' and element == 'Member':
                axis_domain = '도메인'
            elif element == 'Axis':
                axis_domain = '축'
            elif element == 'Member' and prev_axis_domain in ('축', '도메인', '멤버'):
                axis_domain = '멤버'
            else:
                axis_domain = None

            axis_flag = 1 if axis_domain == '축' else 0

            if axis_domain is None:
                group_id = None
            elif axis_flag == 1:
                group_id = 1 if prev_group_id is None else prev_group_id + 1
            else:
                group_id = prev_group_id

            if axis_domain is None:
                axis_name = None
            elif prev_group_id is None or group_id != prev_group_id:
                axis_name = row.get('Name', '') if axis_domain == '축' else ''
            else:
                axis_name = row.get('Name', '') if axis_domain == '축' else prev_axis_name

            key = f"{axis_name}-{row.get('Name', '')}" if axis_domain is not None and axis_name else None

            rows[orig_idx].update({
                '축_도메인': axis_domain,
                'Axis_flag': axis_flag,
                'Axis_Name': axis_name,
                'GroupID':   group_id,
                'KEY_axis':  key,
            })

            prev_element     = element
            prev_axis_domain = axis_domain
            prev_group_id    = group_id
            prev_axis_name   = axis_name


def _parse_basic_info(df, data: TaxonomyXlsxData):
    """메인 버전: '법인명 : 회사명' 한 셀 형식 AND 인접 셀 형식 모두 처리"""
    for i, row in df.iterrows():
        for j, cell in enumerate(row):
            s = _safe(cell)

            if '법인명' in s or '회사명' in s:
                # "법인명 : 인탑스 주식회사" 형식 (한 셀)
                if ':' in s:
                    v = s.split(':', 1)[1].strip()
                    if v:
                        data.company_name = v
                        continue
                # label | value 인접 셀 형식
                for k in range(j + 1, min(j + 4, len(row))):
                    v = _safe(row.iloc[k])
                    if v:
                        data.company_name = v
                        break

            if not data.report_date and ('문서작성일' in s or '회계기간종료일' in s):
                # "문서작성일 : 2024-12-31" 형식 (한 셀)
                if ':' in s:
                    v = s.split(':', 1)[1].strip()
                    if v:
                        data.report_date = v
                        continue
                # label | value 인접 셀 형식
                for k in range(j + 1, min(j + 4, len(row))):
                    v = _safe(row.iloc[k])
                    if v:
                        data.report_date = v
                        break
