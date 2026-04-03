import io
import os
import uuid
import warnings
import pandas as pd
from flask import Flask, render_template, jsonify, request, session

warnings.filterwarnings('ignore')

app = Flask(__name__)
app.config['JSON_SORT_KEYS'] = False
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.secret_key = os.environ.get('SECRET_KEY', 'esg-analysis-dev-key-change-in-prod')

# ── 상수 (변경 없음) ──────────────────────────────────────────────────────────

GRADE_ORDER = {'A+': 8, 'A': 7, 'B+': 6, 'B': 5, 'C+': 4, 'C': 3, 'D+': 2, 'D': 1, '-': 0}

CATEGORY_INFO = {
    '노동관행':               {'q_start': 2002, 'q_end': 2009, 'color': '#4e79a7'},
    '직장 내 안전보건':        {'q_start': 2010, 'q_end': 2015, 'color': '#f28e2b'},
    '인권':                  {'q_start': 2016, 'q_end': 2020, 'color': '#e15759'},
    '공정운영관행':            {'q_start': 2021, 'q_end': 2029, 'color': '#76b7b2'},
    '지속가능한 소비':          {'q_start': 2030, 'q_end': 2035, 'color': '#59a14f'},
    '정보보호 및 개인정보보호':  {'q_start': 2036, 'q_end': 2041, 'color': '#edc948'},
    '지역사회 참여 및 개발':    {'q_start': 2042, 'q_end': 2045, 'color': '#b07aa1'},
    '이해관계자 소통':          {'q_start': 2046, 'q_end': 2046, 'color': '#ff9da7'},
}

CORE_KPI = [
    {'name': '고용 및 근로조건',               'q_num': 2005, 'max': 3,
     'answers': {'0': '근로자 자발적 이직률 미공개', '3': '근로자 자발적 이직률 공개'}},
    {'name': '노사관계',                       'q_num': 2007, 'max': 3,
     'answers': {'0': '노사간 주기적 소통 채널 확인 불가', '3': '노사간 주기적 소통 채널 운영'}},
    {'name': '직장 내 보건 및 안전 위험 관리',   'q_num': 2012, 'max': 7,
     'answers': {'0': '안전보건 개선 활동 미공개', '2': '안전보건경영 활동 공개',
                 '3': '안전보건 관련 위험 요소 도출', '5': '안전보건 위험 완화 조치 수립',
                 '7': '안전보건 위험 완화 조치의 효과성 평가'}},
    {'name': '인력 개발 및 지원',               'q_num': 2009, 'max': 5,
     'answers': {'0': '교육 프로그램 확인 불가', '3': '교육 프로그램 성과 측정',
                 '5': '교육 프로그램 효과성 측정'}},
    {'name': '인권 위험 관리',                  'q_num': 2018, 'max': 7,
     'answers': {'0': '인권경영 활동 미공개', '2': '인권경영 활동 공개',
                 '3': '인권 관련 위험 요소 도출', '5': '인권 관련 위험 완화 조치 수립',
                 '7': '인권 관련 위험 완화 조치의 효과성 평가'}},
    {'name': '공정운영관행',                    'q_num': 2022, 'max': 7,
     'answers': {'0': '불공정거래, 부정경쟁 예방 활동 미공개', '2': '불공정거래, 부정경쟁 예방 활동 공개',
                 '3': '불공정거래, 부정경쟁 관련 위험 평가 실시', '5': '불공정거래, 부정경쟁 관련 위험 완화 조치 수립',
                 '7': '불공정거래, 부정경쟁 관련 위험 완화 조치의 효과성 평가'}},
    {'name': '공급망 위험 관리',                'q_num': 2028, 'max': 7,
     'answers': {'0': '공급망 위험 관리 활동 미공개', '2': '신규 협력사 대상 공급망 위험 관리 실시',
                 '3': '정기적 공급망 위험 관리 실시', '5': '정기적 공급망 위험 관리에 따라 인센티브 부여',
                 '7': '정기적 공급망 위험 관리에 따라 관리 조치 실시'}},
    {'name': '소비자 권익 침해 위험 관리',       'q_num': 2032, 'max': 7,
     'answers': {'0': '소비자 권익 침해 예방 활동 미공개', '2': '소비자 권익 침해 예방 활동 공개',
                 '3': '소비자 권익 침해 위험 평가 실시', '5': '소비자 권익 침해 위험 완화 조치 수립',
                 '7': '소비자 권익 침해 위험 완화 조치의 효과성 평가'}},
    {'name': '정보보호 및 개인정보보호 위험 관리', 'q_num': 2038, 'max': 7,
     'answers': {'0': '정보보호 및 개인정보보호 활동 미공개', '2': '정보보호 및 개인정보보호 활동 공개',
                 '3': '정보보호 및 개인정보보호 관련 위험 평가 실시',
                 '5': '정보보호 및 개인정보보호 관련 위험 완화 조치 수립',
                 '7': '정보보호 및 개인정보보호 관련 위험 완화 조치의 효과성 평가'}},
    {'name': '정보보호 및 개인정보보호 투자 현황', 'q_num': 2040, 'max': 5,
     'answers': {'0': '확인불가 또는 전체 IT 예산 대비 5% 미만',
                 '3': '전체 IT 예산 대비 5% 이상',
                 '5': '전체 IT예산 대비 7% 이상 또는 정보보호 투자 우수기업'}},
    {'name': '지역사회 위험관리',                'q_num': 2044, 'max': 7,
     'answers': {'0': '지역사회에 미칠 수 있는 부정적 영향 미공개',
                 '3': '지역사회에 미칠 수 있는 부정적 영향을 구체적으로 공개',
                 '5': '지역사회에 미칠 수 있는 부정적 영향 완화 조치 수립',
                 '7': '지역사회에 미칠 수 있는 부정적 영향 완화 조치의 효과성 평가'}},
    {'name': '지역사회 참여 성과 관리',           'q_num': 2045, 'max': 3,
     'answers': {'0': '지역사회 상생활동 확인 불가', '2': '지역사회 상생 활동 공개',
                 '3': '지역사회 상생 활동의 성과 공개'}},
    {'name': '이해관계자 소통',                  'q_num': 2046, 'max': 3,
     'answers': {'0': '정보 공개 채널 없음', '2': '주요 비재무정보 공개',
                 '3': '비재무정보에 대한 제3자 검증 실시'}},
]

# ── 세션별 데이터 저장소 ────────────────────────────────────────────────────────
SESSION_DATA = {}  # { session_id: { 'years': {...}, 'companies': {...} } }


# ── 핵심 파싱 함수 (로직 100% 동일, 파일경로 → BytesIO로만 변경) ──────────────────

def parse_uploaded_files(grade_bytes, eval_bytes, adv_bytes):
    result = {'years': {}, 'companies': {}}

    # --- 1. ESG 등급 ---
    df_g = pd.read_excel(io.BytesIO(grade_bytes))
    df_g = df_g[df_g['기업코드'].notna() & df_g['기업명'].notna()]
    for _, row in df_g.iterrows():
        code = str(int(row['기업코드']))
        name = str(row['기업명']).strip()
        if code not in result['companies']:
            result['companies'][code] = {'name': name, 'code': code}

    result['years']['2025'] = {}
    result['years']['2025']['grades'] = {}
    for _, row in df_g.iterrows():
        if pd.isna(row['기업코드']):
            continue
        code = str(int(row['기업코드']))
        result['years']['2025']['grades'][code] = {
            'env':        str(row['환경등급'])     if not pd.isna(row.get('환경등급', None))     else '-',
            'social':     str(row['사회등급'])     if not pd.isna(row.get('사회등급', None))     else '-',
            'gov':        str(row['지배구조등급'])  if not pd.isna(row.get('지배구조등급', None))  else '-',
            'total':      str(row['전체등급'])     if not pd.isna(row.get('전체등급', None))     else '-',
            'adj_e':      str(row['E 등급조정'])   if not pd.isna(row.get('E 등급조정', None))   else None,
            'adj_s':      str(row['S 등급조정'])   if not pd.isna(row.get('S 등급조정', None))   else None,
            'adj_g':      str(row['G 등급조정'])   if not pd.isna(row.get('G 등급조정', None))   else None,
            'adj_esg':    str(row['ESG 등급조정']) if not pd.isna(row.get('ESG 등급조정', None)) else None,
            'reason_e':   str(row['E 조정사유'])   if not pd.isna(row.get('E 조정사유', None))   else None,
            'reason_s':   str(row['S 조정사유'])   if not pd.isna(row.get('S 조정사유', None))   else None,
            'reason_g':   str(row['G 조정사유'])   if not pd.isna(row.get('G 조정사유', None))   else None,
            'reason_esg': str(row['ESG 조정사유']) if not pd.isna(row.get('ESG 조정사유', None)) else None,
        }

    # --- 2. 사회 평가결과 ---
    df_raw = pd.read_excel(io.BytesIO(eval_bytes), header=None)

    q_col_map = {}
    q_text_map = {}
    cat_col_map = {}
    current_cat = None
    for col_idx in range(len(df_raw.columns)):
        cat_val   = df_raw.iloc[0, col_idx]
        q_num_val = df_raw.iloc[1, col_idx]
        q_text_val = df_raw.iloc[2, col_idx]
        if not pd.isna(cat_val) and str(cat_val) not in ('nan', ''):
            current_cat = str(cat_val).strip()
        if not pd.isna(q_num_val):
            try:
                q_num = int(float(q_num_val))
                q_col_map[q_num] = col_idx
                q_text_map[q_num] = str(q_text_val).strip() if not pd.isna(q_text_val) else ''
                if current_cat:
                    cat_col_map.setdefault(current_cat, []).append(col_idx)
            except (ValueError, TypeError):
                pass

    result['years']['2025']['q_col_map']  = q_col_map
    result['years']['2025']['q_text_map'] = q_text_map
    result['years']['2025']['cat_col_map'] = cat_col_map

    score_cols = {}
    for col_idx in range(len(df_raw.columns)):
        val = df_raw.iloc[0, col_idx]
        if not pd.isna(val):
            label = str(val).strip()
            if label in ('기본평가', '감점', '기본평가(백분위)', '감점(백분위)', '기본평가(환산)', '감점(환산)'):
                score_cols[label] = col_idx

    social_scores = {}
    for row_idx in range(3, len(df_raw)):
        name_val = df_raw.iloc[row_idx, 0]
        code_val = df_raw.iloc[row_idx, 1]
        if pd.isna(name_val) or pd.isna(code_val):
            continue
        try:
            code = str(int(float(code_val)))
        except (ValueError, TypeError):
            continue

        name   = str(name_val).strip()
        sector = str(df_raw.iloc[row_idx, 3]).strip() if not pd.isna(df_raw.iloc[row_idx, 3]) else ''
        group  = str(df_raw.iloc[row_idx, 4]).strip() if not pd.isna(df_raw.iloc[row_idx, 4]) else ''

        q_scores = {}
        for q_num, col_idx in q_col_map.items():
            raw_val = df_raw.iloc[row_idx, col_idx]
            if pd.isna(raw_val) or str(raw_val).strip() == '-':
                q_scores[q_num] = None
            else:
                try:
                    q_scores[q_num] = float(raw_val)
                except (ValueError, TypeError):
                    q_scores[q_num] = None

        summary_scores = {}
        for label, col_idx in score_cols.items():
            raw_val = df_raw.iloc[row_idx, col_idx]
            try:
                summary_scores[label] = float(raw_val) if not pd.isna(raw_val) else None
            except (ValueError, TypeError):
                summary_scores[label] = None

        social_scores[code] = {
            'name': name, 'sector': sector, 'group': group,
            'q_scores': q_scores, 'summary': summary_scores,
        }
        if code not in result['companies']:
            result['companies'][code] = {'name': name, 'code': code}

    result['years']['2025']['social_scores'] = social_scores

    # --- 3. 심화평가 (HTML .xls) ---
    try:
        dfs = pd.read_html(io.BytesIO(adv_bytes))
        df_adv = dfs[0]

        adv_by_company  = {}
        total_deductions = {}
        for _, row in df_adv.iterrows():
            try:
                code = str(int(float(row['기업코드'])))
            except (ValueError, TypeError):
                continue
            name      = str(row['기업명']).strip()
            q_text    = str(row['문항']).strip()            if not pd.isna(row.get('문항'))          else ''
            incident  = str(row['사건']).strip()            if not pd.isna(row.get('사건'))          else ''
            details   = str(row['과정 또는 결과']).strip()  if not pd.isna(row.get('과정 또는 결과')) else ''
            deduction = float(row['감점'])                  if not pd.isna(row.get('감점'))          else 0.0

            if code not in adv_by_company:
                adv_by_company[code]   = {'name': name, 'items': [], 'total': 0.0}
                total_deductions[code] = 0.0

            adv_by_company[code]['items'].append({
                'code':      str(row.get('심화평가코드', '')),
                'question':  q_text,
                'incident':  incident,
                'details':   details,
                'deduction': deduction,
            })
            adv_by_company[code]['total'] += deduction
            total_deductions[code] = adv_by_company[code]['total']

        deduction_values = [v for v in total_deductions.values() if v > 0]
        avg_deduction    = sum(deduction_values) / len(deduction_values) if deduction_values else 0

        for code in adv_by_company:
            total = adv_by_company[code]['total']
            if total == 0:
                risk = '위험 없음'
            elif total > avg_deduction:
                risk = '위험 높음'
            elif abs(total - avg_deduction) <= avg_deduction * 0.05:
                risk = '위험'
            else:
                risk = '위험 유의'
            adv_by_company[code]['risk']          = risk
            adv_by_company[code]['avg_deduction'] = avg_deduction

        result['years']['2025']['adv_eval']      = adv_by_company
        result['years']['2025']['avg_deduction'] = avg_deduction

    except Exception as e:
        print(f'심화평가 파싱 오류: {e}')
        result['years']['2025']['adv_eval']      = {}
        result['years']['2025']['avg_deduction'] = 0

    return result


# ── 세션 헬퍼 ─────────────────────────────────────────────────────────────────

def get_session_data():
    sid = session.get('sid')
    if not sid or sid not in SESSION_DATA:
        return None, ('세션이 없습니다. 파일을 먼저 업로드해주세요.', 401)
    return SESSION_DATA[sid], None


# ── 라우트 ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def api_upload():
    keys = ['grade_file', 'eval_file', 'adv_file']
    for key in keys:
        if key not in request.files or request.files[key].filename == '':
            return jsonify({'error': f'파일 누락: {key}'}), 400

    # 세션 수 제한 (메모리 보호)
    if len(SESSION_DATA) > 100:
        oldest = next(iter(SESSION_DATA))
        SESSION_DATA.pop(oldest, None)

    try:
        grade_bytes = request.files['grade_file'].read()
        eval_bytes  = request.files['eval_file'].read()
        adv_bytes   = request.files['adv_file'].read()
        data = parse_uploaded_files(grade_bytes, eval_bytes, adv_bytes)
    except Exception as e:
        return jsonify({'error': f'파일 파싱 오류: {str(e)}'}), 422

    sid = str(uuid.uuid4())
    SESSION_DATA[sid] = data
    session['sid'] = sid

    return jsonify({
        'session_id':    sid,
        'company_count': len(data['companies']),
    })


@app.route('/api/companies')
def api_companies():
    data, err = get_session_data()
    if err:
        return jsonify({'error': err[0]}), err[1]

    year       = '2025'
    grades_data = data['years'].get(year, {}).get('grades', {})
    social_data = data['years'].get(year, {}).get('social_scores', {})

    companies = []
    for code, info in data['companies'].items():
        grade_info  = grades_data.get(code, {})
        social_info = social_data.get(code, {})
        companies.append({
            'code':         code,
            'name':         info['name'],
            'social_grade': grade_info.get('social', '-'),
            'total_grade':  grade_info.get('total', '-'),
            'sector':       social_info.get('sector', ''),
        })

    companies.sort(key=lambda x: x['name'])
    return jsonify(companies)


@app.route('/api/company/<code>')
def api_company(code):
    data, err = get_session_data()
    if err:
        return jsonify({'error': err[0]}), err[1]

    year      = request.args.get('year', '2025')
    year_data = data['years'].get(year, {})
    companies = data['companies']

    if code not in companies:
        return jsonify({'error': '기업을 찾을 수 없습니다'}), 404

    result = {'code': code, 'name': companies[code]['name'], 'year': year}

    grades = year_data.get('grades', {}).get(code, {})
    result['grades'] = grades

    social      = year_data.get('social_scores', {}).get(code, {})
    q_text_map  = year_data.get('q_text_map', {})
    q_scores    = social.get('q_scores', {})

    kpi_list = []
    for kpi in CORE_KPI:
        q_num     = kpi['q_num']
        score     = q_scores.get(q_num)
        score_key = str(int(score)) if score is not None else None
        label     = kpi['answers'].get(score_key, '해당없음') if score_key is not None else '해당없음'
        kpi_list.append({
            'name':      kpi['name'],
            'q_num':     q_num,
            'score':     score,
            'label':     label,
            'max_score': kpi['max'],
        })
    result['kpi'] = kpi_list

    cat_list = []
    for cat_name, cat_info in CATEGORY_INFO.items():
        questions = []
        total     = 0.0
        answered  = 0
        for q_num in range(cat_info['q_start'], cat_info['q_end'] + 1):
            score = q_scores.get(q_num)
            questions.append({
                'q_num': q_num,
                'text':  q_text_map.get(q_num, f'문항 {q_num}'),
                'score': score,
            })
            if score is not None:
                total    += score
                answered += 1
        cat_list.append({
            'name':            cat_name,
            'questions':       questions,
            'total':           total,
            'answered':        answered,
            'total_questions': len(questions),
            'color':           cat_info['color'],
        })
    result['categories']     = cat_list
    result['social_summary'] = social.get('summary', {})
    result['sector']         = social.get('sector', '')
    result['group']          = social.get('group', '')

    adv_data = year_data.get('adv_eval', {}).get(code, {'items': [], 'total': 0.0, 'risk': '위험 없음'})
    adv_data['avg_deduction'] = year_data.get('avg_deduction', 0)
    result['adv_eval'] = adv_data

    return jsonify(result)


@app.route('/api/grade_distribution')
def api_grade_distribution():
    data, err = get_session_data()
    if err:
        return jsonify({'error': err[0]}), err[1]

    year   = request.args.get('year', '2025')
    grades = data['years'].get(year, {}).get('grades', {})
    dist   = {}
    for code, g in grades.items():
        for field in ('social', 'env', 'gov', 'total'):
            grade = g.get(field, '-')
            if grade and grade not in ('-', 'nan'):
                dist.setdefault(field, {})
                dist[field][grade] = dist[field].get(grade, 0) + 1

    return jsonify(dist)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
