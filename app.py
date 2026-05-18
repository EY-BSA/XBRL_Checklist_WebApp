"""
app.py  — XBRL 체크리스트 웹 애플리케이션 (Flask)
입력: IxD 편집기 '구조내려받기' .xlsx 파일 전용 
"""
import io, os, json, traceback
from flask import Flask, request, jsonify, render_template, send_file

from taxonomy_xlsx_parser import parse_taxonomy_xlsx
from checklist_engine import run_all_checks, get_summary
from standard_taxonomy import (StandardTaxonomy,
                               enrich_axis_domain_check, enrich_dart_negate_check)

app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024   # 100 MB

# 단순 메모리 캐시 (프로세스 내 마지막 결과 보관)
_cache: dict = {}

# 데이터 파일 경로
_AXIS_DOMAIN_CHECK_PATH = os.path.join(os.path.dirname(__file__), 'data', 'Axis_Domain_Check.xlsx')
_DART_NEGATE_CHECK_PATH = os.path.join(os.path.dirname(__file__), 'data', 'DART_Negate_Check.xlsx')


def _load_std():
    """표준 택사노미 로드.
    Axis_Domain_Check.xlsx → DART_Negate_Check.xlsx 순으로 보강.
    """
    std = StandardTaxonomy()

    if os.path.exists(_AXIS_DOMAIN_CHECK_PATH):
        try:
            enrich_axis_domain_check(std, _AXIS_DOMAIN_CHECK_PATH)
        except Exception:
            pass

    if os.path.exists(_DART_NEGATE_CHECK_PATH):
        try:
            enrich_dart_negate_check(std, _DART_NEGATE_CHECK_PATH)
        except Exception:
            pass

    return std


# ─── 직렬화 ──────────────────────────────────────────────────────────────────

def _serialize(data, results, summary) -> dict:
    """CheckResult 목록을 JSON-직렬화 가능한 dict로 변환."""
    out = {
        'company':      data.company_name or '알 수 없음',
        'report_date':  data.report_date  or '',
        'entity_id':    data.entity_id    or '',
        'parse_errors': data.errors,
        'summary': {
            'total_checks':       summary['total_checks'],
            'total_issues':       summary['total_issues'],
            'checks_with_issues': summary['checks_with_issues'],
        },
        'categories':  {},
        'checks_flat': [],
    }

    for cat_name, cat_data in summary['categories'].items():
        out['categories'][cat_name] = {
            'total_issues': cat_data['total_issues'],
            'checks': [],
        }
        for chk in cat_data['checks']:
            issues_out = []
            for iss in chk.issues[:500]:          # 시트당 최대 500건
                iss_dict = {
                    'role_uri':        iss.role_uri,
                    'role_code':       iss.role_code,
                    'role_name_ko':    iss.role_name_ko,
                    'role_name_en':    iss.role_name_en,
                    'is_consolidated': iss.is_consolidated,
                    'element_name':    iss.element_name,
                    'label_ko':        iss.label_ko,
                    'label_en':        iss.label_en,
                    'label_role':      iss.label_role,
                    'prefix':          iss.prefix,
                    'data_type':       iss.data_type,
                    'balance':         iss.balance,
                    'period':          iss.period,
                    'gubn':            iss.gubn,
                    'depth':           iss.depth,
                    'parent_name':     iss.parent_name,
                    'parent_label_ko': iss.parent_label_ko,
                    'parent_gubn':     iss.parent_gubn,
                    'reason':          iss.reason,
                    'table_name_ko':   iss.table_name_ko,
                    'client_negate':   iss.client_negate,
                    'dart_negate':     iss.dart_negate,
                }
                issues_out.append(iss_dict)

            chk_dict = {
                'id':           chk.check_id,
                'title':        chk.title,
                'description':  chk.description,
                'category':     chk.category,
                'sheet':        chk.sheet,
                'issue_count':  chk.issue_count,
                'consol_count': chk.consol_count,
                'sep_count':    chk.sep_count,
                'passed':       chk.passed,
                'issues':       issues_out,
            }
            out['categories'][cat_name]['checks'].append(chk_dict)
            out['checks_flat'].append({
                'id':          chk.check_id,
                'title':       chk.title,
                'issue_count': chk.issue_count,
                'passed':      chk.passed,
            })

    return out


# ─── 라우트 ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/check', methods=['POST'])
def api_check():
    """구조내려받기 xlsx → 29개 체크리스트 실행"""
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400

    f = request.files['file']
    if not f.filename.lower().endswith('.xlsx'):
        return jsonify({
            'error': '.xlsx 파일만 지원합니다. '
                     'IxD 편집기에서 [구조내려받기]로 내보낸 파일을 업로드해 주세요.'
        }), 400

    try:
        xlsx_bytes = f.read()
        data       = parse_taxonomy_xlsx(xlsx_bytes)
        std        = _load_std()
        results    = run_all_checks(data, std)
        summary    = get_summary(results)
        out        = _serialize(data, results, summary)
        _cache['last'] = out
        return jsonify(out)
    except Exception as e:
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500


@app.route('/api/export', methods=['GET'])
def api_export():
    """마지막 체크 결과를 XBRL CoE Checklist_Result 템플릿 형식으로 내보내기"""
    cached = _cache.get('last')
    if not cached:
        return jsonify({'error': '먼저 파일을 업로드하세요.'}), 400

    try:
        import openpyxl

        TEMPLATE_PATH = os.path.join(os.path.dirname(__file__),
                                     'template', 'XBRL_CoE_Checklist_Result.xlsx')
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({'error': '템플릿 파일을 찾을 수 없습니다: template/XBRL_CoE_Checklist_Result.xlsx'}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)

        def _role_def(iss: dict) -> str:
            """Role Definition 컬럼 값: [코드] 한글명 | 영문명"""
            rc = iss.get('role_code', '')
            ko = iss.get('role_name_ko', '')
            en = iss.get('role_name_en', '')
            if rc and ko:
                return f"[{rc}] {ko} | {en}" if en else f"[{rc}] {ko}"
            return ko or rc

        def _table_name(iss: dict) -> str:
            """TABLE NAME 컬럼 값: table_name_ko 우선, 없으면 [코드] 한글명"""
            tko = iss.get('table_name_ko', '')
            if tko:
                return tko
            rc = iss.get('role_code', '')
            ko = iss.get('role_name_ko', '')
            if rc and ko:
                return f"[{rc}] {ko}"
            return ko or rc

        def _write_standard_row(ws, ri: int, iss: dict):
            """표준 컬럼(B~J) 에 이슈 1행 기록"""
            ws.cell(ri, 2, _role_def(iss))           # B: Role Definition
            ws.cell(ri, 3, _table_name(iss))          # C: TABLE NAME
            ws.cell(ri, 4, iss.get('prefix', ''))   # D: Prefix
            ws.cell(ri, 5, iss.get('element_name', ''))  # E: Name
            ws.cell(ri, 6, iss.get('label_ko', '')) # F: Label(KO)
            ws.cell(ri, 7, iss.get('label_en', '')) # G: Label(EN)
            ws.cell(ri, 8, iss.get('label_role', ''))    # H: Label Role
            ws.cell(ri, 9, iss.get('data_type', ''))     # I: DataType
            ws.cell(ri, 10, iss.get('period', ''))  # J: Period

        def _clear_data_rows(ws):
            """데이터 행(14행~) 값 초기화"""
            if ws.max_row < 14:
                return
            for row in ws.iter_rows(min_row=14, max_row=ws.max_row):
                for cell in row:
                    cell.value = None

        # ── 각 체크 결과를 해당 시트에 기록 ──────────────────────────────
        for cat_name, cat_data in cached['categories'].items():
            for chk in cat_data['checks']:
                sheet_name = chk.get('sheet', f"Checklist_{chk['id']}")
                sheet_name = sheet_name[:31]

                if sheet_name not in wb.sheetnames:
                    # 템플릿에 없는 시트(예: Checklist_7-2)는 건너뜀
                    continue

                ws = wb[sheet_name]
                _clear_data_rows(ws)

                if not chk['issues']:
                    continue

                chk_id = chk['id']

                for ri, iss in enumerate(chk['issues'], 14):
                    if chk_id == '5-6':
                        # 5-6: B=Role Definition, C=단위미표시 (Checklist_5-6 헤더: 단위표시/미표시)
                        ws.cell(ri, 2, _table_name(iss))
                        ws.cell(ri, 3, '단위미표시')

                    elif chk_id == '7-1':
                        # 7-1: Negate — 표준(B~J) + K=DART_Negate, L=Client_Negate
                        _write_standard_row(ws, ri, iss)
                        ws.cell(ri, 11, iss.get('dart_negate', ''))    # K: DART_Negate
                        ws.cell(ri, 12, iss.get('client_negate', ''))  # L: Client_Negate

                    else:
                        # 나머지 모든 시트: 표준 컬럼(B~J)
                        _write_standard_row(ws, ri, iss)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        company = cached.get('company') or ''
        if not company or company == '알 수 없음':
            company = cached.get('entity_id') or 'XBRL'
        rdate   = (cached.get('report_date') or '').replace('/', '').replace('-', '')
        name_parts = [p for p in [company, rdate] if p]
        fname   = f"XBRL_CoE_Checklist_Result_{'_'.join(name_parts)}.xlsx"
        return send_file(
            buf, as_attachment=True,
            download_name=fname,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, port=port)
