import datetime
import io
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.styles import Font

# ================== 상수 ==================

DATA_DIR = Path(__file__).parent / 'data'
# 과제 명세 파일명 우선, 실습 폴더에 실제로 있는 파일명도 허용
PROJECT_FILES = ['SB_projects_demo.csv', 'SB_project_demo.csv']
RESOURCE_FILES = ['SB_resources_demo.csv', 'SB_resource_demo.csv']

REQUIRED_COLS = ['의뢰번호', '고객사', '제품명', '모델명', '시험분야', '적용규격',
                 '담당엔지니어', '접수일', '완료예정일', '진행상태', '시료상태', '메모']
PROGRESS_ORDER = ['접수', '자료확인', '시험진행', '성적서검토', '보완요청', '완료']
FILTER_COLS = ['시험분야', '적용규격', '담당엔지니어', '진행상태', '시료상태']
TABLE_COLS = ['의뢰번호', '고객사', '제품명', '모델명', '시험분야', '적용규격', '담당엔지니어',
              '접수일', '완료예정일', '잔여일수', '일정상태', '진행상태', '시료상태', '메모']
INITIAL_FIELDS = ['EMC', '무선(RF)', '인증서류', 'SAR']

DEFAULT_BASE_DATE = datetime.date(2026, 10, 23)

DONE = '완료'
DELAY = '지연 후보(일정 확인 필요)'
DUE_TODAY = '오늘 마감'
PLANNED = '진행 예정'
CHECK = '일정 확인 필요'
SCHEDULE_ORDER = [DELAY, DUE_TODAY, PLANNED, CHECK, DONE]
EMPTY_MSG = '조건에 맞는 인증 프로젝트가 없습니다'

SCREEN_NOTICE = '교육용 가상 데이터 기준 — 실제 업무 시스템 아님'
EXCEL_NOTICE = '교육용 검토 파일 — 공식 보고서 아님'
RESOURCE_NOTICE = '현황 조회용이며 예약·시험 재개를 승인하는 기능이 아님'

RESOURCE_STATUS_ORDER = ['사용 가능', '사용 중', '점검 중']
RESOURCE_COLORS = {'사용 가능': '#1E7B4F', '사용 중': '#1F5FBF', '점검 중': '#B26A00'}

NAVY = '#0B2545'
BLUE = '#1F5FBF'
BLUE_SCALE = ['#0B2545', '#13315C', '#1F5FBF', '#4A86D9', '#8DB3E8', '#C9DBF5']


# ================== 데이터 읽기 ==================

def find_data_file(names):
    """data 폴더에서 후보 파일명 중 처음 존재하는 파일 경로를 반환 (없으면 None)."""
    for name in names:
        path = DATA_DIR / name
        if path.exists():
            return path
    return None


def read_csv(source):
    """로컬 경로 또는 업로드 파일을 문자열 DataFrame으로 읽음. UTF-8 → CP949 순으로 시도."""
    raw = source.read_bytes() if isinstance(source, Path) else source.getvalue()
    for enc in ('utf-8-sig', 'cp949'):
        try:
            return pd.read_csv(io.BytesIO(raw), dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError('CSV 인코딩을 읽을 수 없습니다 (UTF-8 또는 CP949로 저장해주세요).')


def init_session_state(projects, source_key):
    """원본 의뢰 데이터와 시험분야 목록을 session_state에 한 번만 복사.
    다른 데이터 파일이 들어오면 새로 초기화한다."""
    if st.session_state.get('source_key') == source_key:
        return
    projects = projects.copy()
    projects[FILTER_COLS] = projects[FILTER_COLS].fillna('(미입력)')
    projects['메모'] = projects['메모'].fillna('')
    # 데이터에만 있는 시험분야도 관리 목록에 포함해 '고아' 분류가 생기지 않게 함
    extra = [f for f in projects['시험분야'].unique() if f not in INITIAL_FIELDS]
    st.session_state.projects = projects
    st.session_state.fields = INITIAL_FIELDS + sorted(extra)
    st.session_state.source_key = source_key


# ================== 일정 상태 ==================

def schedule_status(progress, remaining):
    """진행상태와 잔여일수로 일정상태를 판단. 위에서부터 먼저 맞는 조건이 적용된다."""
    # 1) 완료 건은 예정일이 지났어도 지연이 아님
    if progress == DONE:
        return DONE
    # 2) 완료예정일이 비었거나 날짜 형식이 아니면 잔여일수를 계산할 수 없음
    if pd.isna(remaining):
        return CHECK
    if remaining < 0:
        return DELAY
    if remaining == 0:
        return DUE_TODAY
    return PLANNED


def add_schedule_columns(df, base_date):
    """잔여일수(완료예정일 - 기준일)와 일정상태 열을 추가한 사본을 반환."""
    df = df.copy()
    due = pd.to_datetime(df['완료예정일'], format='%Y-%m-%d', errors='coerce')
    df['잔여일수'] = (due - pd.Timestamp(base_date)).dt.days.astype('Int64')
    df['일정상태'] = [schedule_status(p, r) for p, r in zip(df['진행상태'], df['잔여일수'])]
    return df


def highlight_delay(row):
    """표에서 지연 후보 행을 빨간색 계열로 칠하는 Styler 함수."""
    style = 'background-color: #FDE2E2; color: #9B1C1C' if row['일정상태'] == DELAY else ''
    return [style] * len(row)


# ================== 집계 (화면·엑셀 공용) ==================

def count_by(df, col, order):
    """col 값별 건수와 비율 표. order에 있는 값은 0건이어도 그 순서대로 포함."""
    counts = df[col].value_counts()
    keys = order + sorted(k for k in counts.index if k not in order)
    result = pd.DataFrame({col: keys, '건수': [int(counts.get(k, 0)) for k in keys]})
    total = result['건수'].sum()
    result['비율'] = (result['건수'] / total).round(3) if total else 0.0
    return result


def engineer_summary(df):
    """담당엔지니어 × 진행상태 건수표에 합계와 지연 후보 건수를 붙인 표."""
    if df.empty:
        return pd.DataFrame(columns=['담당엔지니어', *PROGRESS_ORDER, '합계', '지연 후보'])
    table = pd.crosstab(df['담당엔지니어'], df['진행상태'])
    table = table.reindex(columns=PROGRESS_ORDER, fill_value=0)
    table['합계'] = table.sum(axis=1)
    table['지연 후보'] = df[df['일정상태'] == DELAY].groupby('담당엔지니어').size()
    table['지연 후보'] = table['지연 후보'].fillna(0).astype(int)
    return table.reset_index()


def apply_filters(df, selections):
    """사이드바에서 선택한 값으로 행을 거른 결과를 반환."""
    mask = pd.Series(True, index=df.index)
    for col, selected in selections.items():
        mask &= df[col].isin(selected)
    return df[mask]


# ================== Excel ==================

def write_sheet(writer, sheet_name, tables):
    """시트 1행에 교육용 안내를 쓰고, 그 아래에 (제목, 표) 목록을 차례로 기록."""
    row = 2
    for title, table in tables:
        # openpyxl은 pd.NA를 못 쓰므로 빈 값은 None으로 바꿔서 기록
        table = table.astype(object).where(table.notna(), None)
        table.to_excel(writer, sheet_name=sheet_name, startrow=row + 1, index=False)
        writer.sheets[sheet_name].cell(row=row + 1, column=1, value=title).font = Font(bold=True)
        row += len(table) + 4
    ws = writer.sheets[sheet_name]
    ws['A1'] = EXCEL_NOTICE
    ws['A1'].font = Font(bold=True, color='C00000')
    for col_cells in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells[1:]) if len(col_cells) > 1 else 8
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(width * 1.6, 8), 40)


def build_excel(df, conditions):
    """프로젝트목록 / 상태별집계 / 담당자별집계 / 분석조건 4개 시트의 xlsx 바이트를 생성."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        write_sheet(writer, '프로젝트목록', [('프로젝트 목록', df[TABLE_COLS])])
        write_sheet(writer, '상태별집계', [
            ('진행상태별 건수', count_by(df, '진행상태', PROGRESS_ORDER)),
            ('일정상태별 건수', count_by(df, '일정상태', SCHEDULE_ORDER)),
        ])
        write_sheet(writer, '담당자별집계', [('담당엔지니어별 건수', engineer_summary(df))])
        write_sheet(writer, '분석조건', [('분석 조건', conditions)])
    return buffer.getvalue()


def analysis_conditions(scope, base_date, selections, row_count):
    """분석조건 시트에 기록할 항목/값 표. selections가 None이면 필터 미적용(전체)."""
    rows = [
        ('다운로드 범위', scope),
        ('다운로드 일시', datetime.datetime.now().strftime('%Y-%m-%d %H:%M')),
        ('기준일', str(base_date)),
        ('건수', f'{row_count}건'),
    ]
    for col in FILTER_COLS:
        if selections is None:
            value = '적용 안 함 (전체)'
        else:
            value = ', '.join(selections[col]) if selections[col] else '(선택 없음)'
        rows.append((f'필터: {col}', value))
    return pd.DataFrame(rows, columns=['항목', '값'])


# ================== 화면 구성 요소 ==================

def apply_style():
    """화이트 배경 + 네이비/블루 포인트 CSS와 상단 안내 배너를 그림."""
    st.markdown(f"""
    <style>
    .stApp {{ background-color: #FFFFFF; color: #1A1A1A; }}
    [data-testid="stSidebar"] {{ background-color: #F3F6FB; }}
    h1, h2, h3 {{ color: {NAVY} !important; }}
    .sb-banner {{
        background: {NAVY}; color: #FFFFFF; padding: 10px 16px; border-radius: 6px;
        font-weight: 600; margin-bottom: 8px;
    }}
    [data-testid="stMetric"] {{
        background: #FFFFFF; border: 1px solid #D6E2F3; border-left: 4px solid {BLUE};
        border-radius: 6px; padding: 10px 14px;
    }}
    [data-testid="stMetricLabel"] p {{ color: {NAVY}; font-weight: 600; }}
    [data-testid="stMetricValue"] {{ color: {BLUE}; }}
    .sb-resource {{
        border: 1px solid #D6E2F3; border-radius: 6px; padding: 12px 14px; background: #FFFFFF;
    }}
    .sb-resource .name {{ color: {NAVY}; font-weight: 700; font-size: 1.05rem; }}
    .sb-resource .badge {{
        display: inline-block; color: #FFFFFF; border-radius: 4px; padding: 2px 8px;
        font-size: 0.85rem; font-weight: 600; margin: 6px 0;
    }}
    .sb-resource .note {{ color: #555555; font-size: 0.85rem; }}
    </style>
    """, unsafe_allow_html=True)
    st.markdown(f'<div class="sb-banner">{SCREEN_NOTICE}</div>', unsafe_allow_html=True)


def metric_row(items):
    """(라벨, 건수) 목록을 한 줄의 요약 카드로 표시."""
    for col, (label, count) in zip(st.columns(len(items)), items):
        col.metric(label, f'{count}건')


def set_flash(kind, message):
    """다음 실행(rerun) 때 보여줄 결과 메시지를 저장."""
    st.session_state.flash = (kind, message)


def show_flash():
    """저장된 결과 메시지를 한 번 보여주고 지움."""
    if 'flash' in st.session_state:
        kind, message = st.session_state.pop('flash')
        getattr(st, kind)(message)


def ordered_options(df, col):
    """필터 선택지. 시험분야는 관리 목록, 진행상태는 정해진 순서, 나머지는 가나다순."""
    values = list(df[col].unique())
    if col == '시험분야':
        return st.session_state.fields + sorted(v for v in values if v not in st.session_state.fields)
    if col == '진행상태':
        return [p for p in PROGRESS_ORDER if p in values] + sorted(v for v in values if v not in PROGRESS_ORDER)
    return sorted(values)


def render_filters(df):
    """사이드바 '필터' 영역: 기준일과 다중 선택 필터를 그리고 선택값을 반환."""
    st.header('필터')
    base_date = st.date_input('기준일', DEFAULT_BASE_DATE)
    selections = {}
    for col in FILTER_COLS:
        options = ordered_options(df, col)
        selections[col] = st.multiselect(col, options, default=options)
    return base_date, selections


def rename_field(old, new):
    """시험분야 이름을 바꾸고 해당 분야의 의뢰 데이터도 함께 갱신."""
    fields = st.session_state.fields
    fields[fields.index(old)] = new
    projects = st.session_state.projects
    projects.loc[projects['시험분야'] == old, '시험분야'] = new


def reclassify_field(source, target):
    """source 분야의 의뢰를 모두 target 분야로 옮기고 옮긴 건수를 반환."""
    projects = st.session_state.projects
    mask = projects['시험분야'] == source
    projects.loc[mask, '시험분야'] = target
    return int(mask.sum())


def render_category_manager():
    """사이드바 '카테고리 관리' 영역: 시험분야 등록 / 이름 변경 / 삭제."""
    st.header('카테고리 관리')
    show_flash()
    fields = st.session_state.fields
    projects = st.session_state.projects
    action = st.radio('작업', ['등록', '이름 변경', '삭제'], horizontal=True, key='category_action')

    if action == '등록':
        with st.form('add_field', clear_on_submit=True):
            name = st.text_input('새 시험분야 이름')
            if st.form_submit_button('등록'):
                name = name.strip()
                if not name:
                    st.error('이름을 입력해주세요.')
                elif name in fields:
                    st.error(f"'{name}'은(는) 이미 있는 시험분야입니다.")
                else:
                    fields.append(name)
                    set_flash('success', f"'{name}' 시험분야를 등록했습니다.")
                    st.rerun()

    elif action == '이름 변경':
        with st.form('rename_field', clear_on_submit=True):
            old = st.selectbox('변경할 시험분야', fields)
            new = st.text_input('새 이름')
            if st.form_submit_button('이름 변경'):
                new = new.strip()
                if not new:
                    st.error('새 이름을 입력해주세요.')
                elif new in fields:
                    st.error(f"'{new}'은(는) 이미 있는 시험분야입니다.")
                else:
                    moved = int((projects['시험분야'] == old).sum())
                    rename_field(old, new)
                    set_flash('success', f"'{old}' → '{new}'로 변경했습니다. 관련 의뢰 {moved}건도 함께 갱신했습니다.")
                    st.rerun()

    else:
        target = st.selectbox('삭제할 시험분야', fields, key='delete_target')
        in_use = int((projects['시험분야'] == target).sum())
        if in_use == 0:
            if st.button('삭제', type='primary'):
                fields.remove(target)
                set_flash('success', f"'{target}' 시험분야를 삭제했습니다.")
                st.rerun()
            return
        # 사용 중인 분야는 바로 삭제하지 않고, 재분류를 먼저 하도록 안내
        st.warning(f"'{target}' 분야는 의뢰 {in_use}건에서 사용 중이라 바로 삭제할 수 없습니다. "
                   "아래에서 다른 분야로 재분류한 뒤 삭제해주세요.")
        others = [f for f in fields if f != target]
        if not others:
            st.info('옮길 다른 시험분야가 없습니다. 먼저 새 분야를 등록해주세요.')
            return
        dest = st.selectbox('재분류할 분야', others, key='reclassify_dest')
        if st.button(f'{in_use}건을 {dest}(으)로 재분류'):
            moved = reclassify_field(target, dest)
            set_flash('success', f"'{target}' 의뢰 {moved}건을 '{dest}'(으)로 재분류했습니다. 이제 삭제할 수 있습니다.")
            st.rerun()


def render_resources():
    """시험 자원 상태 카드. CSV 값을 그대로 보여줄 뿐 상태를 바꾸지 않는다
    ('점검 중'을 자동으로 '사용 가능'으로 바꾸는 로직 없음)."""
    st.subheader('시험 자원 현황')
    path = find_data_file(RESOURCE_FILES)
    if path is None:
        st.info(f'data 폴더에 {RESOURCE_FILES[0]} 파일이 없어 자원 현황을 표시하지 않습니다.')
        return
    resources = read_csv(path).fillna('')
    status_counts = resources['상태'].value_counts()
    metric_row([(s, int(status_counts.get(s, 0))) for s in RESOURCE_STATUS_ORDER])

    for col, (_, res) in zip(st.columns(len(resources)), resources.iterrows()):
        color = RESOURCE_COLORS.get(res['상태'], '#6B7280')
        col.markdown(
            f'<div class="sb-resource"><div class="name">{res["자원"]}</div>'
            f'<span class="badge" style="background:{color}">{res["상태"]}</span>'
            f'<div class="note">{res["비고"]}</div></div>',
            unsafe_allow_html=True,
        )
    st.caption(f'※ {RESOURCE_NOTICE}')


def render_charts(df):
    """시험분야별 건수 막대그래프와 진행상태 비율 도넛 차트."""
    left, right = st.columns(2)
    with left:
        st.subheader('시험분야별 건수')
        field_counts = df['시험분야'].value_counts().rename_axis('시험분야').reset_index(name='건수')
        bar = alt.Chart(field_counts).mark_bar(color=BLUE).encode(
            x=alt.X('시험분야:N', sort='-y', axis=alt.Axis(labelAngle=0)),
            y=alt.Y('건수:Q', axis=alt.Axis(tickMinStep=1)),
            tooltip=['시험분야', '건수'],
        )
        st.altair_chart(bar, width='stretch')
    with right:
        st.subheader('진행상태 비율')
        ratio = count_by(df, '진행상태', PROGRESS_ORDER)
        ratio = ratio[ratio['건수'] > 0]
        donut = alt.Chart(ratio).mark_arc(innerRadius=60).encode(
            theta='건수:Q',
            color=alt.Color('진행상태:N', sort=PROGRESS_ORDER,
                            scale=alt.Scale(domain=PROGRESS_ORDER, range=BLUE_SCALE)),
            tooltip=['진행상태', '건수', alt.Tooltip('비율:Q', format='.0%')],
        )
        st.altair_chart(donut, width='stretch')


def render_downloads(all_df, filtered, base_date, selections):
    """전체 목록 / 현재 필터 결과를 각각 xlsx로 내려받는 버튼 (파일명에 오늘 날짜)."""
    today = datetime.date.today().strftime('%Y%m%d')
    left, right = st.columns(2)
    left.download_button(
        f'전체 목록 다운로드 ({len(all_df)}건)',
        build_excel(all_df, analysis_conditions('전체 목록', base_date, None, len(all_df))),
        file_name=f'standardbank_project_status_{today}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        width='stretch',
    )
    right.download_button(
        f'현재 필터 결과 다운로드 ({len(filtered)}건)',
        build_excel(filtered, analysis_conditions('현재 필터 결과', base_date, selections, len(filtered))),
        file_name=f'standardbank_project_status_filtered_{today}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        width='stretch',
    )


def render_detail(df):
    """의뢰번호를 선택하면 상세 정보와 담당자 메모를 보여주고, 메모를 수정·저장."""
    st.subheader('의뢰 상세조회')
    req_id = st.selectbox('의뢰번호', df['의뢰번호'].tolist())
    row = df[df['의뢰번호'] == req_id].iloc[0]
    remaining = '-' if pd.isna(row['잔여일수']) else f"{row['잔여일수']}일"

    left, right = st.columns(2)
    left.markdown(
        f"**고객사** {row['고객사']}  \n"
        f"**제품 · 모델** {row['제품명']} / {row['모델명']}  \n"
        f"**시험분야** {row['시험분야']}  \n"
        f"**적용규격** {row['적용규격']}"
    )
    right.markdown(
        f"**담당엔지니어** {row['담당엔지니어']}  \n"
        f"**시료상태** {row['시료상태']}  \n"
        f"**완료예정일** {row['완료예정일'] if pd.notna(row['완료예정일']) else '(미입력)'}  \n"
        f"**잔여일수 / 일정상태** {remaining} / {row['일정상태']}"
    )

    show_flash_key = f'memo_saved_{req_id}'
    memo = st.text_area('담당자 메모', row['메모'], key=f'memo_input_{req_id}')
    if st.button('메모 저장'):
        projects = st.session_state.projects
        projects.loc[projects['의뢰번호'] == req_id, '메모'] = memo
        st.session_state[show_flash_key] = True
        st.rerun()
    if st.session_state.pop(show_flash_key, False):
        st.success(f'{req_id} 메모를 저장했습니다.')


# ================== 메인 ==================

st.set_page_config(page_title='SB 인증 프로젝트 대시보드', layout='wide')
apply_style()
st.title('인증 프로젝트 접수 · 진행현황 대시보드')
st.caption('ⓘ 시험분야·메모 변경 내용은 이 브라우저 세션(session_state)에만 저장되며, '
           '새로고침하거나 세션이 종료되면 사라질 수 있습니다. 원본 CSV는 수정되지 않습니다.')

data_path = find_data_file(PROJECT_FILES)
if data_path is None:
    st.warning(f'data 폴더에서 {PROJECT_FILES[0]} 파일을 찾을 수 없습니다. '
               '아래에서 CSV 파일을 직접 올려주세요.')
    source = st.file_uploader('인증 프로젝트 CSV 업로드', type='csv')
    if source is None:
        st.stop()
    source_key = f'upload:{source.name}:{source.size}'
else:
    source = data_path
    source_key = f'file:{data_path}'

if st.session_state.get('source_key') != source_key:
    try:
        raw_projects = read_csv(source)
    except Exception as e:
        st.error(f'CSV 파일을 읽지 못했습니다: {e}')
        st.stop()
    missing = [c for c in REQUIRED_COLS if c not in raw_projects.columns]
    if missing:
        st.error(f"필수 컬럼이 없습니다: {', '.join(missing)}")
        st.stop()
    init_session_state(raw_projects, source_key)

with st.sidebar:
    base_date, selections = render_filters(st.session_state.projects)
    st.divider()
    render_category_manager()

df = add_schedule_columns(st.session_state.projects, base_date)
filtered = apply_filters(df, selections)

st.caption(f'기준일 {base_date} · 전체 {len(df)}건 중 {len(filtered)}건 표시 · '
           f'데이터: {getattr(source, "name", source)}')

render_resources()
st.divider()

if filtered.empty:
    # 0건이면 카드·차트·표·상세 계산을 하지 않음 (전체 목록 다운로드는 가능)
    st.info(EMPTY_MSG)
else:
    st.subheader('진행상태 요약')
    progress = count_by(filtered, '진행상태', PROGRESS_ORDER)
    metric_row([('전체', len(filtered))] + list(zip(progress['진행상태'], progress['건수'])))

    st.subheader('일정 요약')
    schedule = count_by(filtered, '일정상태', SCHEDULE_ORDER).set_index('일정상태')['건수']
    metric_row([(s.split('(')[0], int(schedule[s])) for s in [DELAY, DUE_TODAY, PLANNED, CHECK]])

    st.divider()
    render_charts(filtered)
    st.divider()

    st.subheader('의뢰 목록')
    st.dataframe(filtered[TABLE_COLS].style.apply(highlight_delay, axis=1),
                 hide_index=True, width='stretch')

st.subheader('Excel 다운로드')
render_downloads(df, filtered, base_date, selections)

if not filtered.empty:
    st.divider()
    render_detail(filtered)
