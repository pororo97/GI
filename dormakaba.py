import base64
import pandas as pd
import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
import dash_bootstrap_components as dbc

# PDFファイルのパス
pdf_files = {
    "23/24決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\dormakaba\dormakaba 23_24決算情報.pdf"
}

# PDFをBase64エンコードする関数
def encode_pdf(file_path):
    try:
        with open(file_path, 'rb') as pdf_file:
            encoded_pdf = base64.b64encode(pdf_file.read()).decode('utf-8')
        return f"data:application/pdf;base64,{encoded_pdf}"
    except FileNotFoundError:
        return None

# 各シートのデータを読み込む
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\dormakaba\dorma.xlsx"
sales_sheet = "Sales"
info_sheet = "Information"
ma_sheet = "MA"
news_sheet = "News"
ebit_sheet = "EBIT"

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)

df_ma = pd.read_excel(file_path, sheet_name=ma_sheet)
df_ma['Year'] = pd.to_datetime(df_ma['Year'], errors='coerce').dt.year

df_news = pd.read_excel(file_path, sheet_name=news_sheet, header=0)
df_news.fillna('', inplace=True)

# Salesデータの読み込み
sales_data = pd.read_excel(file_path, sheet_name=sales_sheet)

# EBITデータの読み込み
df_ebit = pd.read_excel(file_path, sheet_name=ebit_sheet)

# 必要な列の定義
columns = ["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1"]
ebit_columns = ["調整前EBIT", "EBITへの調整", "調整(のれん償却)", "調整(リストラIAC)", "調整後EBIT"]
ebit_columns_jpy = ["調整前EBIT(JPY)", "EBITへの調整(JPY)", "調整(のれん償却)(JPY)", "調整(リストラIAC)(JPY)", "調整後EBIT(JPY)"]

# 現地通貨データの抽出
access_solutions_local = sales_data[sales_data["Sales"] == "Access Solutions"][columns].values[0]
key_wall_solutions_local = sales_data[sales_data["Sales"] == "Key & Wall Solutions"][columns].values[0]
eliminations_local = sales_data[sales_data["Sales"] == "Eliminations"][columns].values[0]
group_local = sales_data[sales_data["Sales"] == "Group"][columns].values[0]

# 日本円換算データの抽出
access_solutions_jpy = sales_data[sales_data["Sales"] == "Access Solutions(JPY)"][columns].values[0]
key_wall_solutions_jpy = sales_data[sales_data["Sales"] == "Key & Wall Solutions(JPY)"][columns].values[0]
eliminations_jpy = sales_data[sales_data["Sales"] == "Eliminations(JPY)"][columns].values[0]
group_jpy = sales_data[sales_data["Sales"] == "Group(JPY)"][columns].values[0]

def generate_info_content(dataframe):
    content = []
    for _, row in dataframe.iterrows():
        content.append(html.P([
            html.Strong(f"{row['テーマ']}: "),  # テーマ列を太字に
            html.Span(row['内容'])
        ], style={'font-size': '10px', 'margin': '0px 0'}))
    return content

def generate_news_content(dataframe):
    content = []
    for _, row in dataframe.iterrows():
        content.append(html.Div([
            html.H4(row['English'], style={'margin': '0'}),
            html.P(row['日本語'], style={'margin': '5px 0'}),
            html.P(row['Date'], style={'margin': '5px 0', 'font-size': '12px', 'color': 'gray'}),
            html.A("続き読む", href=row['Link'], target="_blank", style={'font-size': '12px', 'color': 'blue'})
        ], style={'background-color': '#B0E0E6', 'padding-top': '0px', 'margin': '5px 0'}))
    return content

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

image_url = "https://img.securityinfowatch.com/files/base/cygnus/siw/image/2017/02/dormakaba.589dceadaa0a6.png?auto=format%2Ccompress&amp;amp;w=640&amp;amp;width=640"

def create_layout(app):
    return html.Div([
    html.Div([
        html.Div(generate_info_content(df_info), style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-bottom': '0px', 'flex': '1'}),
        html.Div([html.Img(src=image_url, style={'width': '80%', 'height': '80%', 'margin-bottom': '0px'})], style={'flex': '1', 'padding': '1px'})
    ], style={'display': 'flex', 'justify-content': 'space-between'}),

    html.Div(html.H2("News", style={
        'text-align': 'center',
        'font-size': '24px',
        'font-weight': 'bold',
        'margin-bottom': '0px',
        'color': 'black',
        'background-color': '#ADD8E6',
        'padding': '10px',
        'border-radius': '5px'
    }), style={'margin': '0px'}),

    html.Div(generate_news_content(df_news), style={'background-color': '#f9f9f9', 'padding': '3px', 'margin': '0px'}),
    

    dcc.RadioItems(
        id='currency-type-ytd',
        options=[
            {'label': '現地通貨(MCHF)', 'value': 'local'},
            {'label': '日本円換算(億円)', 'value': 'jpy'}
        ],
        value='local',
        labelStyle={'display': 'inline-block', 'margin-right': '10px'},
        style={'margin-bottom': '20px'}
    ),

    # Salesデータのグラフを追加
    dcc.Graph(id='sales-chart', style={'height': '400px'}),

    # EBITデータのグラフを追加
    dcc.Graph(id='ebit-chart', style={'height': '400px'}),
    
    html.Div([
        dcc.Dropdown(
            id='year-filter',
            options=[{'label': str(year), 'value': year} for year in df_ma['Year'].dropna().unique()],
            multi=True,
            placeholder="Select Year"
        ),
    ], style={'margin-bottom': '10px', 'width': '50%', 'display': 'inline-block', 'background-color': '#acecf2'}),
    
    html.Div([
        dcc.Graph(id='choropleth-map', style={'background-color': '#acecf2', 'width': '50%', 'display': 'inline-block', 'height': '400px'}),
        dash_table.DataTable(
            id='ma-table',
            columns=[{"name": i, "id": i} for i in df_ma.columns],
            page_size=10,
            style_table={'background-color': '#acecf2', 'overflowX': 'auto', 'maxHeight': '400px', 'overflowY': 'auto'},
            style_cell={'textAlign': 'left', 'font-size': '8px', 'whiteSpace': 'normal', 'background-color': '#E2EDFF', 'height': 'auto'},
            style_data_conditional=[{'if': {'column_id': c}, 'width': '100px'} for c in df_ma.columns],
            page_action='native',
            style_as_list_view=True,
        )
    ], style={'display': 'flex'}),

    

    html.Div([
        html.Label("決算情報を選択してください:", style={'font-weight': 'bold'}),
        dcc.Dropdown(
            id='pdf-selector',
            options=[{'label': name, 'value': name} for name in pdf_files.keys()],
            value="23/24決算情報",  # デフォルト値
            style={'width': '50%'}
        )
    ], style={'margin-bottom': '20px'}),

    # PDFファイル名を表示
    html.Div(id='pdf-filename-display', style={'margin-bottom': '10px', 'font-size': '16px', 'font-weight': 'bold'}),

    # PDF表示用のiframe
    html.Iframe(
        id='pdf-viewer',
        src="",
        style={'width': '80%', 'height': '600px', 'border': 'none', 'margin': '0 auto', 'display': 'block'}
    ),
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
])

@app.callback(
    Output('ma-table', 'data'),
    [Input('year-filter', 'value')]
)
def update_ma_table(selected_years):
    filtered_ma_df = df_ma.copy()
    if selected_years:
        filtered_ma_df = filtered_ma_df[filtered_ma_df['Year'].isin(selected_years)]
    return filtered_ma_df.to_dict('records')

@app.callback(
    Output('choropleth-map', 'figure'),
    [Input('year-filter', 'value')]
)
def update_map(selected_years):
    filtered_df = df_ma.copy()
    if selected_years:
        filtered_df = filtered_df[filtered_df['Year'].isin(selected_years)]
    if filtered_df.empty:
        return go.Figure()
    fig = px.choropleth(
        filtered_df,
        locations="Country",
        locationmode="country names",
        color="HC",
        hover_name="Company",
        hover_data={
            "Year": True,
            "Industry": True,
            "Category": True,
            "Strategy": True,
            "HC": True,
            "Country": True,
            "City": True,
        },
        title="M&A Map",
        projection="natural earth"
    )
    fig.update_geos(showcoastlines=True, showland=True)
    fig.update_layout(geo=dict(bgcolor='white'), margin={"r": 0, "t": 0, "l": 0, "b": 0}, coloraxis_showscale=False)
    return fig

@app.callback(
    Output("sales-chart", "figure"),
    [Input("currency-type-ytd", "value")]
)
def update_chart(currency):
    # 根据选择的货币类型切换数据
    if currency == "local":
        access = access_solutions_local
        key_wall = key_wall_solutions_local
        elim = eliminations_local
        group_data = group_local
        yaxis_title = "Value (MCHF)"
    elif currency == "jpy":
        access = access_solutions_jpy
        key_wall = key_wall_solutions_jpy
        elim = eliminations_jpy
        group_data = group_jpy
        yaxis_title = "Value (億円)"

    # 创建堆积柱状图
    fig = go.Figure()

    # 年ごとにデータを整理
    half_years = [f"{col}" for col in columns]  # 半期のラベル

    # 堆积柱状图
    fig.add_trace(go.Bar(
        x=half_years,
        y=access,
        name="Access Solutions",
        marker_color="pink",
        offsetgroup=0))

    fig.add_trace(go.Bar(
        x=half_years,
        y=key_wall,
        name="Key & Wall Solutions",
        marker_color="orange",
        base=access,  # 前のトレースの上に積み上げる
        offsetgroup=0))

    # 折線図（Group）
    fig.add_trace(go.Scatter(
        x=half_years,
        y=group_data,
        name="Group",
        mode="lines+markers",
        marker=dict(color="red", size=10),
        line=dict(width=1)))

    # 更新布局
    fig.update_layout(
        barmode="relative",  # グループモード
        title="Sales",
        xaxis_title="Period",
        yaxis_title=yaxis_title,
        legend_title="Legend",
        template="plotly_white"
    )

    return fig

@app.callback(
    Output("ebit-chart", "figure"),
    [Input("currency-type-ytd", "value")]
)
def update_ebit_chart(currency):
    # 現地通貨または日本円換算に基づいてEBITデータを選択
    if currency == "local":
        ebit_data = df_ebit.loc[df_ebit["EBIT"].isin(
            ["EBITDA", "調整(のれん償却)", "調整(リストラIAC)"]),
            ["EBIT", "2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"]
        ]
        adjusted_ebit = df_ebit.loc[df_ebit['EBIT'] == "調整後EBIT", ["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"]].values[0]
        base_ebit = df_ebit.loc[df_ebit['EBIT'] == "調整前EBIT", ["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"]].values[0]
        yaxis_title = "EBIT (MCHF)"
    elif currency == "jpy":
        ebit_data = df_ebit.loc[df_ebit["EBIT"].isin(
            ["EBITDA(JPY)", "調整(のれん償却)(JPY)", "調整(リストラIAC)(JPY)"]),
            ["EBIT", "2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"]
        ]
        adjusted_ebit = df_ebit.loc[df_ebit['EBIT'] == "調整後EBIT(JPY)", ["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"]].values[0]
        base_ebit = df_ebit.loc[df_ebit['EBIT'] == "調整前EBIT(JPY)", ["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"]].values[0]
        yaxis_title = "EBIT (億円)"

    # EBITの堆積柱状図を作成
    fig = go.Figure()

    # 堆積柱状図
    for row in ebit_data.iterrows():
        fig.add_trace(go.Bar(
            x=["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"],
            y=row[1].values[1:],  # 各行のデータを取得
            name=row[1].values[0],  # 行名を取得
            offsetgroup=0
        ))

    # 折れ線図（調整後EBIT）
    fig.add_trace(go.Scatter(
        x=["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"],
        y=adjusted_ebit,
        name="調整後EBIT",
        mode="lines+markers",
        marker=dict(color="green", size=10),
        line=dict(width=2)
    ))

    # 折れ線図（調整前EBIT）
    fig.add_trace(go.Scatter(
        x=["2019H2", "2020H1", "2020H2", "2021H1", "2021H2", "2022H1", "2022H2", "2023H1", "2023H2", "2024H1", "2024H2"],
        y=base_ebit,
        name="調整前EBIT",
        mode="lines+markers",
        marker=dict(color="blue", size=10),
        line=dict(width=2)
    ))

    # 更新レイアウト
    fig.update_layout(
        title="EBIT",
        xaxis_title="Period",
        yaxis_title=yaxis_title,
        legend_title="Legend",
        template="plotly_white"
    )

    return fig

@app.callback(
    [Output('pdf-viewer', 'src'), Output('pdf-filename-display', 'children')],
    [Input('pdf-selector', 'value')]
)
def update_pdf_viewer(selected_pdf_name):
    selected_pdf_path = pdf_files.get(selected_pdf_name)
    if selected_pdf_path:
        encoded_pdf = encode_pdf(selected_pdf_path)
        if encoded_pdf:
            # デバッグ用にエンコード結果をログに出力
            print("Encoded PDF:", encoded_pdf[:100])  # 最初の100文字だけ表示
            return encoded_pdf, f"表示中のPDF: {selected_pdf_name}"
        else:
            return "", "エラー: PDFファイルが見つかりません。"
    else:
        return "", "エラー: 選択されたPDFオプションが無効です。"

if __name__ == '__main__':
    app.layout = create_layout(app)  # レイアウトを設定
    app.run_server(debug=True)