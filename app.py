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


# ── 핵심 파싱 함수 ────────────────────────────────────────────────────────────

def _safe_str(val, default='-'):
    return str(val).strip() if not pd.isna(val) else default

def _safe_none(val):
    return str(val).strip() if not pd.isna(val) else None

def parse_year_data(year, grade_bytes, eval_bytes, adv_bytes, result):
    """단일 연도 파일 3개를 파싱해 result dict에 추가한다."""

    result['years'][year] = {}

    # --- 1. ESG 등급 (벡터화) ---
    df_g = pd.read_excel(io.BytesIO(grade_bytes))
    df_g = df_g[df_g['기업코드'].notna() & df_g['기업명'].notna()].copy()
    df_g['_code'] = df_g['기업코드'].apply(lambda x: str(int(x)))

    for row in df_g[['_code', '기업명']].itertuples(index=False):
        if row._code not in result['companies']:
            result['companies'][row._code] = {'name': str(row[1]).strip(), 'code': row._code}

    def _col(df, name, default='-'):
        return df[name].apply(lambda x: _safe_str(x, default)) if name in df.columns else pd.Series(default, index=df.index)
    def _col_none(df, name):
        return df[name].apply(_safe_none) if name in df.columns else pd.Series(None, index=df.index)

    grades_df = pd.DataFrame({
        'code':       df_g['_code'],
        'env':        _col(df_g, '환경등급'),
        'social':     _col(df_g, '사회등급'),
        'gov':        _col(df_g, '지배구조등급'),
        'total':      _col(df_g, '전체등급'),
        'adj_e':      _col_none(df_g, 'E 등급조정'),
        'adj_s':      _col_none(df_g, 'S 등급조정'),
        'adj_g':      _col_none(df_g, 'G 등급조정'),
        'adj_esg':    _col_none(df_g, 'ESG 등급조정'),
        'reason_e':   _col_none(df_g, 'E 조정사유'),
        'reason_s':   _col_none(df_g, 'S 조정사유'),
        'reason_g':   _col_none(df_g, 'G 조정사유'),
        'reason_esg': _col_none(df_g, 'ESG 조정사유'),
    })
    result['years'][year]['grades'] = {
        r['code']: {k: v for k, v in r.items() if k != 'code'}
        for r in grades_df.to_dict('records')
    }

    # --- 2. 사회 평가결과 (벡터화) ---
    df_raw = pd.read_excel(io.BytesIO(eval_bytes), header=None)

    # 헤더 3행에서 컬럼 매핑 구성
    row0 = df_raw.iloc[0]
    row1 = df_raw.iloc[1]
    row2 = df_raw.iloc[2]

    q_col_map  = {}
    q_text_map = {}
    cat_col_map = {}
    score_col_map = {}
    current_cat = None
    SCORE_LABELS = {'기본평가', '감점', '기본평가(백분위)', '감점(백분위)', '기본평가(환산)', '감점(환산)'}

    for ci in range(len(df_raw.columns)):
        cat_val = row0.iloc[ci]
        if not pd.isna(cat_val) and str(cat_val).strip() not in ('nan', ''):
            current_cat = str(cat_val).strip()
            if str(cat_val).strip() in SCORE_LABELS:
                score_col_map[str(cat_val).strip()] = ci
                continue
        q_val = row1.iloc[ci]
        if not pd.isna(q_val):
            try:
                q_num = int(float(q_val))
                q_col_map[q_num] = ci
                q_text_map[q_num] = str(row2.iloc[ci]).strip() if not pd.isna(row2.iloc[ci]) else ''
                if current_cat:
                    cat_col_map.setdefault(current_cat, []).append(ci)
            except (ValueError, TypeError):
                pass

    result['years'][year]['q_col_map']   = q_col_map
    result['years'][year]['q_text_map']  = q_text_map
    result['years'][year]['cat_col_map'] = cat_col_map

    # 데이터 행만 슬라이스 (행 3부터)
    df_data = df_raw.iloc[3:].copy()
    df_data = df_data[df_data.iloc[:, 0].notna() & df_data.iloc[:, 1].notna()]

    # 기업 코드 벡터 변환
    def to_code(x):
        try: return str(int(float(x)))
        except: return None
    df_data['_code'] = df_data.iloc[:, 1].apply(to_code)
    df_data = df_data[df_data['_code'].notna()]

    # 문항 점수 컬럼 일괄 추출
    q_cols  = list(q_col_map.keys())
    q_cidxs = [q_col_map[q] for q in q_cols]

    # 숫자 변환: '-' → NaN → None
    score_block = df_data.iloc[:, q_cidxs].copy()
    score_block = score_block.replace('-', float('nan'))
    score_block = score_block.apply(pd.to_numeric, errors='coerce')

    # summary 컬럼 일괄 추출
    sum_labels = list(score_col_map.keys())
    sum_cidxs  = [score_col_map[l] for l in sum_labels]
    sum_block  = df_data.iloc[:, sum_cidxs].copy().apply(pd.to_numeric, errors='coerce')

    social_scores = {}
    codes   = df_data['_code'].tolist()
    names   = df_data.iloc[:, 0].astype(str).str.strip().tolist()
    sectors = df_data.iloc[:, 3].fillna('').astype(str).str.strip().tolist()
    groups  = df_data.iloc[:, 4].fillna('').astype(str).str.strip().tolist()

    score_arr = score_block.values  # numpy array - 빠름
    sum_arr   = sum_block.values

    for i, code in enumerate(codes):
        q_scores = {q_cols[j]: (None if pd.isna(score_arr[i, j]) else float(score_arr[i, j]))
                    for j in range(len(q_cols))}
        summary  = {sum_labels[j]: (None if pd.isna(sum_arr[i, j]) else float(sum_arr[i, j]))
                    for j in range(len(sum_labels))}
        social_scores[code] = {
            'name': names[i], 'sector': sectors[i], 'group': groups[i],
            'q_scores': q_scores, 'summary': summary,
        }
        if code not in result['companies']:
            result['companies'][code] = {'name': names[i], 'code': code}

    result['years'][year]['social_scores'] = social_scores

    # --- 3. 심화평가 (벡터화) ---
    try:
        df_adv = pd.read_html(io.BytesIO(adv_bytes))[0]

        df_adv['_code'] = df_adv['기업코드'].apply(to_code)
        df_adv = df_adv[df_adv['_code'].notna()]
        df_adv['감점'] = pd.to_numeric(df_adv['감점'], errors='coerce').fillna(0.0)
        if '과정 또는 결과' not in df_adv.columns:
            df_adv['과정 또는 결과'] = ''
        if '심화평가코드' not in df_adv.columns:
            df_adv['심화평가코드'] = ''

        adv_by_company = {}
        for row in df_adv.itertuples(index=False):
            code = row._code
            if code not in adv_by_company:
                adv_by_company[code] = {'name': str(row[df_adv.columns.get_loc('기업명')]), 'items': [], 'total': 0.0}
            ded = float(row[df_adv.columns.get_loc('감점')])
            adv_by_company[code]['items'].append({
                'code':      str(row[df_adv.columns.get_loc('심화평가코드')]),
                'question':  str(row[df_adv.columns.get_loc('문항')]),
                'incident':  str(row[df_adv.columns.get_loc('사건')]),
                'details':   str(row[df_adv.columns.get_loc('과정 또는 결과')]),
                'deduction': ded,
            })
            adv_by_company[code]['total'] += ded

        deduction_values = [v['total'] for v in adv_by_company.values() if v['total'] > 0]
        avg_deduction = sum(deduction_values) / len(deduction_values) if deduction_values else 0

        for code, entry in adv_by_company.items():
            t = entry['total']
            if t == 0:                                    risk = '위험 없음'
            elif t > avg_deduction:                       risk = '위험 높음'
            elif abs(t - avg_deduction) <= avg_deduction * 0.05: risk = '위험'
            else:                                         risk = '위험 유의'
            entry['risk'] = risk
            entry['avg_deduction'] = avg_deduction

        result['years'][year]['adv_eval']      = adv_by_company
        result['years'][year]['avg_deduction'] = avg_deduction

    except Exception as e:
        print(f'심화평가 파싱 오류: {e}')
        result['years'][year]['adv_eval']      = {}
        result['years'][year]['avg_deduction'] = 0


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
    # 세션 수 제한 (메모리 보호)
    if len(SESSION_DATA) > 100:
        oldest = next(iter(SESSION_DATA))
        SESSION_DATA.pop(oldest, None)

    result = {'years': {}, 'companies': {}}
    years_parsed = []

    for i in range(10):
        year_key  = f'year_{i}'
        grade_key = f'grade_file_{i}'
        eval_key  = f'eval_file_{i}'
        adv_key   = f'adv_file_{i}'

        if year_key not in request.form:
            break

        grade_f = request.files.get(grade_key)
        eval_f  = request.files.get(eval_key)
        adv_f   = request.files.get(adv_key)

        if not grade_f or not eval_f or not adv_f:
            continue
        if grade_f.filename == '' or eval_f.filename == '' or adv_f.filename == '':
            continue

        year = request.form[year_key]
        try:
            parse_year_data(year, grade_f.read(), eval_f.read(), adv_f.read(), result)
            years_parsed.append(year)
        except Exception as e:
            return jsonify({'error': f'{year}년 파일 파싱 오류: {str(e)}'}), 422

    if not years_parsed:
        return jsonify({'error': '업로드된 파일이 없습니다. 연도별로 파일 3개를 모두 선택해주세요.'}), 400

    sid = str(uuid.uuid4())
    SESSION_DATA[sid] = result
    session['sid'] = sid

    return jsonify({
        'session_id':    sid,
        'company_count': len(result['companies']),
        'years':         sorted(years_parsed),
    })


@app.route('/api/companies')
def api_companies():
    data, err = get_session_data()
    if err:
        return jsonify({'error': err[0]}), err[1]

    # 가장 최근 연도 기준으로 목록 표시
    year        = max(data['years'].keys()) if data['years'] else '2025'
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

    # ── 연도별 비교 데이터 ──
    sorted_years = sorted(data['years'].keys())
    all_years = {}
    for yr in sorted_years:
        yr_data    = data['years'][yr]
        yr_grades  = yr_data.get('grades', {}).get(code, {})
        yr_social  = yr_data.get('social_scores', {}).get(code, {})
        yr_q       = yr_social.get('q_scores', {})
        yr_adv     = yr_data.get('adv_eval', {}).get(code, {})

        yr_kpi = []
        for kpi in CORE_KPI:
            s = yr_q.get(kpi['q_num'])
            sk = str(int(s)) if s is not None else None
            yr_kpi.append({
                'name':      kpi['name'],
                'score':     s,
                'max_score': kpi['max'],
                'label':     kpi['answers'].get(sk, '해당없음') if sk is not None else '해당없음',
            })

        yr_cats = []
        for cat_name, cat_info in CATEGORY_INFO.items():
            total = sum(
                yr_q.get(q, 0) or 0
                for q in range(cat_info['q_start'], cat_info['q_end'] + 1)
                if yr_q.get(q) is not None
            )
            yr_cats.append({'name': cat_name, 'total': total, 'color': cat_info['color']})

        all_years[yr] = {
            'grades':      yr_grades,
            'summary':     yr_social.get('summary', {}),
            'kpi':         yr_kpi,
            'categories':  yr_cats,
            'adv_total':   yr_adv.get('total', 0),
            'adv_risk':    yr_adv.get('risk', '위험 없음'),
        }

    result['all_years']       = all_years
    result['available_years'] = sorted_years

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
