import dash
from dash import dcc, html, Input, Output, dash_table
import pandas as pd
import plotly.express as px

# 데이터 읽어오기
file_path = 'MES9월마감데이터.xlsx'
sheet_name = '9월 마감내역서 통합'

df = pd.read_excel(file_path, sheet_name=sheet_name)

app = dash.Dash(__name__, suppress_callback_exceptions=True)


###############################
"""
요구사항

1. 9월 공정 마감 현황에서, 총합계가 보여지게 하기.
- 총합계는 필터링이 걸어지는 업체명의 공급가액 모두 합산
- 합산 결과는 Bold로 나오게 하기.

2. 업체명별 마감 상태에서, 0(마감)의 총합계, X(마감X)의 총합계를 추가
또한 업체명별 마감 상태 테이블 오른쪽에는, Stacked Bar테이블로, 업체명에 대한, O마감, X마감 Bar차트로 그리기

3. 공정 필터에 따라, 업체명별 마감 상태 아래에, 업체명을 행으로, Case를 열로 하고,
값을 Case별 갯수를 센, 테이블을 만들고, 오른쪽에, 업체명에 따른, Case를 꺾은선 그래프로 표현하기.

4. 로트번호 검색하기에서, 로트번호를 사용자가 입력시에 나온 테이블에 대해, 각각의 Case를
해당 테이블 아래에, ex) "20241002-200025"의 에러코드는 1이며, 주문원장에 없습니다. 이렇게 나오도록 해줘.

Excel테이블에 내가 Case라는 컬럼이름으로 0 ~ 14까지 에러코드를 숫자로 입력해놨어.
0 ~ 14 의 에러코드는 의미를 가지는데,
로트번호 검색하기 앞에, 에러코드 살펴보기 버튼을 누르면, 0~14까지의 에러코드가 무엇인지 알려주는 테이블이 열리게 했으면 좋겠어.
그리고, 에러코드 살펴보기 버튼을 다시한번 누르면, 해당 테이블이 사라지게 하는거지.
"""


###########################

app.layout = html.Div([
    html.H1("다산팩 9월 마감 대시보드", style={"align-items": 'justify-center'}),
    # TODO: 9월 공정 마감 현황 테이블 + 그래프
    html.Div([
        html.Div([
            html.Label("공정"),
            dcc.Dropdown(
                id='process-filter',
                options=[{'label': p, 'value': p}
                         for p in df['공정'].unique()] + [{'label': 'All', 'value': 'All'}],
                value='All',
                clearable=False,
                className='dropdown'
            ),
            html.H3("9월 공정 마감 현황"),
            dash_table.DataTable(
                id='filtered-table',
                columns=[{"name": "업체명", "id": "업체명"},
                         {"name": "공급가액", "id": "공급가액", 'type': 'numeric',
                             'format': {'specifier': ','}}],
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'center', 'font-size': '16px'},
                style_header={'fontWeight': 'bold', 'text-align': 'center'},
                page_size=10
            ),
            html.H3(id='total-sum', style={'fontWeight': 'bold'}),
        ], style={'width': '40%', 'display': 'inline-block', 'vertical-align': 'top', 'padding': '0 20px'}),

        html.Div([
            dcc.Graph(id='bar-chart'),
        ], style={'width': '55%', 'display': 'inline-block', 'padding-left': '5%'}),
    ], style={'display': 'flex', 'aligh-items': 'flex-start', 'padding': '20px 0'}),  # Ensuring spacing around the section

    # TODO: 업체명별 마감 상태 테이블과 그래프 섹션
    html.Div([
        html.Div([
            html.H3("업체명별 마감 상태"),
            dash_table.DataTable(
                id='status-table',
                columns=[{"name": "업체명", "id": "업체명"},
                         {"name": "O (마감)", "id": "O"},
                         {"name": "X (마감X)", "id": "X"}],
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'center', 'font-size': '16px'},
                style_header={'fontWeight': 'bold', 'text-align': 'center'},
                page_size=10,
                page_action='native',
                page_current=0,
                css=[{
                    'selector': '.current-page',
                    'rule': 'display: none;'
                }]
            ),
            html.H3(id='status-total',
                    style={'fontWeight': 'bold'})
        ], style={'width': '40%', 'display': 'inline-block', 'vertical-align': 'top', 'padding': '0 20px'}),

        # 그래프를 테이블 오른쪽에 배치
        html.Div([
            dcc.Graph(id='status-bar-chart')
        ], style={'width': '55%', 'padding-left': '5%'})
    ], style={'display': 'flex', 'align-items': 'flex-start', 'padding': '20px 0'}),
    # TODO: 업체별, CASE별 갯수 세기, 행을 업체별로, 열을 CASE갯수로, 꺾은선 OR 물방울 그래프 그리기


    # TODO: 에러 테이블 버튼 클릭시, 에러 테이블 뜨게 하기.


    # TODO: 필터로 회사 선택 후 =>  로트 번호를 타이핑하면 실시간으로 검색 기능
    # TODO: 이후, 해당 결과값 테이블의 로트번호 ; 에러코드 ; 에러코드 명으로 인해 마감 작업을 실패하였습니다. 라고 출력되게 하기.
    # 로트번호 검색 기능
    html.Div([
        html.Label("로트번호 검색하기"),
        dcc.Input(id='lot-search', type='text', placeholder="로트번호를 입력하세요."),
        html.Button('에러코드 살펴보기', id='error-code-button', n_clicks=0),
        html.Div(id='error-code-table', style={'display': 'none'}),
        html.Div(id='search-result')
    ], style={'padding': '20px 0'})  # Adding padding to space out the sections
])


@ app.callback(
    [Output('bar-chart', 'figure'),
     Output('filtered-table', 'data'),
     Output('total-sum', 'children')],
    Input('process-filter', 'value')
)
def update_bar_and_table(selected_process):
    filtered_df = df if selected_process == 'All' else df[df['공정']
                                                          == selected_process]
    filtered_df = filtered_df[filtered_df['마감여부'] == 'O']

    # 공급가액 열을 숫자로 변환, 변환할 수 없는 값은 0으로 처리
    filtered_df['공급가액'] = pd.to_numeric(
        filtered_df['공급가액'], errors='coerce').fillna(0)

    # 총합계 계산
    total_sum = filtered_df['공급가액'].sum()

    # Bar chart 데이터
    agg_df = filtered_df.groupby('업체명', as_index=False)['공급가액'].sum()
    fig = px.bar(agg_df, x='업체명', y='공급가액',
                 title='9월 공정 마감 현황',
                 color='공급가액',
                 color_continuous_scale=px.colors.sequential.Blues)
    fig.update_layout(
        yaxis_tickprefix="₩",
        yaxis_tickformat=",",
        yaxis_title="공급가액",
        xaxis_title="업체명",
        title_x=0.5,  # 타이틀 중앙 정렬
        title_font=dict(size=20, color='black'),  # 타이틀 글꼴 설정
        plot_bgcolor='white',  # 배경색을 흰색으로 설정
        paper_bgcolor='white',
        showlegend=False
    )

    # Table 데이터: 업체명으로 그룹화하고 공급가액 합산
    agg_df['공급가액'] = agg_df['공급가액'].apply(lambda x: f"{x:,}")
    table_data = agg_df.to_dict('records')

    # 총합계를 bold로 표시
    total_sum_display = html.H4(f"총합계: ₩{total_sum:,.0f}", style={
        'font-weight': 'bold'})

    return fig, table_data, total_sum_display


@ app.callback(
    [Output('status-table', 'data'),
     Output('status-total', 'children'),
     Output('status-bar-chart', 'figure')],
    Input('process-filter', 'value')
)
def update_status_table(selected_process):
    filtered_df = df if selected_process == 'All' else df[df['공정']
                                                          == selected_process]
    status_count = filtered_df.groupby(
        ['업체명', '마감여부']).size().unstack(fill_value=0).reset_index()

    status_count = status_count.rename(columns={'O': 'O', 'X': 'X'})

    # 마감여부가 없는 경우 0으로 대체
    status_count['O'] = status_count.get('O', 0)
    status_count['X'] = status_count.get('X', 0)

    # O와 X 총합계 계산
    total_o = status_count['O'].sum()
    total_x = status_count['X'].sum()

    # 테이블에 표시할 데이터
    table_data = status_count.to_dict('records')

    # 총합계 표시
    total_sum_display = f"O (마감): {total_o:,} / X (마감X): {total_x:,}"

    # Stacked Bar chart 생성
    fig = px.bar(status_count, x='업체명', y=['O', 'X'],
                 labels={'value': '수량', 'variable': '마감상태'},
                 title='업체명별 마감 상태',
                 barmode='stack')
    fig.update_layout(
        plot_bgcolor='white',
        paper_bgcolor='white',
        yaxis_tickformat=",",
        title_x=0.5,
        title_font=dict(size=20, color='black')
    )

    return table_data, total_sum_display, fig


@ app.callback(
    Output('error-code-table', 'style'),
    Input('error-code-button', 'n_clicks')
)
def toggle_error_table(n_clicks):
    if n_clicks % 2 == 0:
        return {'display': 'none'}
    return {'display': 'block'}


@ app.callback(
    Output('search-result', 'children'),
    Input('lot-search', 'value')
)
def search_lot(lot_no):
    if not lot_no:
        return ""
    result = df[df['LOT NO'].str.contains(lot_no, na=False)]
    if result.empty:
        return "No results found."
    return html.Div([
        dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in result.columns],
            data=result.to_dict('records'),
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'center', 'font-size': '16px'}
        )
    ])


if __name__ == '__main__':
    app.run_server(debug=True)
