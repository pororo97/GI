import base64
import pandas as pd
import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
import base64
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
from dash import html, dcc


layout = html.Div([
    html.H1("Assa App"),
    html.P("Assa")
])

# PDFファイルのパス
pdf_files = {
    "2024ーQ2決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Assa Abloy Q2決算について.pdf",
    "2024ーQ3決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Assa Abloy Q3.pdf",
    "M&A情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\M&A.pdf"
}

# PDFをBase64エンコードする関数
def encode_pdf(file_path):
    try:
        with open(file_path, 'rb') as pdf_file:
            encoded_pdf = base64.b64encode(pdf_file.read()).decode('utf-8')
        return f"data:application/pdf;base64,{encoded_pdf}"
    except FileNotFoundError:
        return None

# ファイルパスとシート名
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\ASSA ABLOY key figures Q3 2024 r.xlsx"
sales_sheet = "Sales"
ebit_sheet = "EBIT"
info_sheet = "Information"
ma_sheet = "M&A"
q_bridge_sheet = "Q Bridge"
news_sheet = "News"
sales_ytd_sheet = "sales YTD"
ebit_ytd_sheet = "EBIT YTD"

# データを読み込む
df_sales = pd.read_excel(file_path, sheet_name=sales_sheet, usecols="A:AO", header=0)
df_sales.columns = df_sales.columns.str.strip()
df_sales.set_index('Sales', inplace=True)


df_ebit = pd.read_excel(file_path, sheet_name=ebit_sheet, usecols="A:AO", header=0)
df_ebit.columns = df_ebit.columns.str.strip()
df_ebit.set_index('EBIT', inplace=True)

df_sales_ytd = pd.read_excel(file_path, sheet_name=sales_ytd_sheet, usecols="A:K", header=0)
df_sales_ytd.columns = df_sales_ytd.columns.str.strip()  # 列名の空白を削除
df_sales_ytd.set_index('Sales', inplace=True)

df_ebit_ytd = pd.read_excel(file_path, sheet_name=ebit_ytd_sheet, usecols="A:K", header=0)
df_ebit_ytd.columns = df_ebit_ytd.columns.str.strip()
df_ebit_ytd.set_index('EBIT', inplace=True)

df_ebit_ytd_percent = pd.read_excel(file_path, sheet_name=ebit_ytd_sheet, usecols="A:K", header=0)  # EBIT%はSales YTDから取る
df_ebit_ytd_percent.columns = df_ebit_ytd_percent.columns.str.strip()
df_ebit_ytd_percent.set_index('EBIT', inplace=True)

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)

df_ma = pd.read_excel(file_path, sheet_name=ma_sheet)
df_ma['Year'] = pd.to_datetime(df_ma['Year'], errors='coerce').dt.year
df_ma['Date'] = pd.to_datetime(df_ma['Date'], errors='coerce').dt.strftime('%Y%m')

df_q_bridge = pd.read_excel(file_path, sheet_name=q_bridge_sheet, usecols="A:AG", header=0)
df_q_bridge.columns = df_q_bridge.columns.str.strip()
df_q_bridge.set_index('Sales Growth', inplace=True)

df_news = pd.read_excel(file_path, sheet_name=news_sheet, header=0)
df_news.fillna('', inplace=True)

# 情報を生成する関数
def generate_info_content(dataframe):
    content = []    
    for _, row in dataframe.iterrows():
        content.append(html.P([
            html.Strong(f"{row['テーマ']}: "),  # テーマ列を太字に
            html.Span(row['内容'])
        ], style={'font-size': '15px','margin': '0px 0'}))
    return content

# News情報を生成する関数
def generate_news_content(dataframe):
    content = []
    for _, row in dataframe.iterrows():
        content.append(html.Div([
            html.H4(row['English'], style={'margin': '0'}),  # 'English'列を使用
            html.P(row['日本語'], style={'margin': '5px 0'}),  # '日本語'列を使用
            html.P(row['Date'], style={'margin': '5px 0', 'font-size': '12px', 'color': 'gray'}),  # 'Date'列を使用
            html.A("続き読む", href=row['Link'], target="_blank", style={'font-size': '12px', 'color': 'blue'})  # 'Link'列を使用
        ], style={'background-color': '#B0E0E6', 'padding-top': '0px', 'margin': '5px 0'}))  # ボーダーとパディングを追加
    return content

# 追加: 画像のURL
image_url ="https://gw-assets.assaabloy.com/is/image/assaabloy/Jpn-Multiple-brand-logo-final2-800x800-1"
# Dashアプリを初期化
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# レイアウトを定義
def create_layout(app):
    return html.Div([
    html.H1("Assa App"),
    html.Div([
        # 情報コンテンツを左側に配置
        html.Div(generate_info_content(df_info), 
                 style={'background-color': '#CDE1FF', 'font-size': '16px','padding': '10px', 'margin-bottom': '0px', 'flex': '3'}),  # 情報内容のDiv

        # 画像を右側に配置
        html.Div([
            html.Img(src=image_url, style={'width': '100%', 'height': '100%', 'margin-bottom': '0px'})
        ], style={'flex': '1', 'padding': '2px'})  # 画像のDiv
    ], style={'display': 'flex', 'justify-content': 'space-between'}),  # フレックスボックスを使用
    
    html.Div(
        html.H2("News", style={
            'text-align': 'center',  # 中央揃え
            'font-size': '24px',  # フォントサイズ
            'font-weight': 'bold',  # 太字
            'margin-bottom': '0px',  # 下部マージン
            'color': 'black',  # 文字色
            'background-color': '#ADD8E6',  # 背景色
            'padding': '10px',  # 内側余白
            'border-radius': '5px'  # 角を丸める
        }),
        style={'margin': '0px'}  # タイトルの外側マージン
    ),
    
    html.Div(generate_news_content(df_news), style={'background-color': '#f9f9f9', 'padding': '3px', 'margin': '0px'}),

    dcc.RadioItems(
        id='assa-currency-type-ytd',
        options=[
            {'label': '現地通貨(MSEK)', 'value': 'local'},
            {'label': '日本円換算(億円)', 'value': 'jpy'}
        ],
        value='local',
        labelStyle={'display': 'inline-block', 'margin-right': '10px'},
        style={'margin-bottom': '20px'}
    ),
    

    dcc.Checklist(
    id='assa-year-checklist',
    options=[{'label': y, 'value': y} for y in df_sales_ytd.columns if pd.notna(y) and y != ''],
    value=list(df_sales_ytd.columns[:10]),style={
        'margin-bottom': '20px',
        'display': 'flex',  # フレックスボックスを使用
        'flex-wrap': 'wrap',  # 必要に応じて折り返し
        'gap': '10px'  # 項目間の間隔を設定
    }
    ),   
    

    html.Div([
    dcc.Graph(id='assa-sales-ytd-chart', style={'width': '33%', 'display': 'inline-block'}),
    dcc.Graph(id='assa-ebit-ytd-chart', style={'width': '33%', 'display': 'inline-block'}),
    dcc.Graph(id='assa-ebit-percent-ytd-chart', style={'width': '33%', 'display': 'inline-block'}),
    ], style={'display': 'flex', 'justify-content': 'space-between'}),


    dcc.RadioItems(
        id='assa-currency-type',
        options=[
            {'label': '現地通貨(MSEK)', 'value': 'local'},
            {'label': '日本円換算(億円)', 'value': 'jpy'}
        ],
        value='local',
        labelStyle={'display': 'inline-block', 'margin-right': '10px'},
        style={'margin-bottom': '20px'}
    ),

    dcc.Checklist(
        id='assa-quarter-checklist',
        options=[{'label': q, 'value': q} for q in df_sales.columns],
        value=df_sales.columns[-15:],  # デフォルトで最初の4つの四半期を選択
        inline=True
    ),

    html.Div([
        dcc.Graph(id='assa-sales-bar-chart', style={'width': '33%', 'display': 'inline-block'}),
        dcc.Graph(id='assa-ebit-bar-chart', style={'width': '33%', 'display': 'inline-block'}),
        dcc.Graph(id='assa-ebit-percent-line-chart', style={'width': '33%', 'display': 'inline-block'})
    ], style={'display': 'flex', 'justify-content': 'space-between'}),

    dcc.RadioItems(
        id='assa-sales-growth-type',
        options=[
            {'label': 'GROUP', 'value': 'group'},
            {'label': 'Entrance', 'value': 'entrance'}
        ],
        value='group',
        labelStyle={'display': 'inline-block', 'margin-right': '10px'},
        style={'margin-bottom': '20px'}
    ),

    dcc.Graph(id='assa-sales-growth-line-chart', style={'width': '100%', 'height': '400px'}),

    html.H2(
        "最新M＆A戦略・売上情報",  # タイトルテキスト
        style={
            'text-align': 'center',  # 中央揃え
            'margin-bottom': '10px',  # 下部余白
            'font-weight': 'bold',  # 太字
            'font-size': '20px',  # フォントサイズ
            'color': '#333333'  # テキスト色
        }
    ),
    
    html.Div([
        dcc.Dropdown(
            id='assa-year-filter',
            options=[{'label': str(year), 'value': year} for year in df_ma['Year'].dropna().unique()],
            multi=True,
            placeholder="Select Year"
        ),
        dcc.Dropdown(
            id='assa-sbu-filter',
            options=[{'label': sbu, 'value': sbu} for sbu in df_ma['SBU'].dropna().unique()],
            multi=True,
            placeholder="Select SBU"
        )
    ], style={'margin-bottom': '10px', 'width': '50%', 'display': 'inline-block','background-color': '#acecf2'}),

    html.Div([
        dcc.Graph(id='assa-choropleth-map', style={'background-color': '#acecf2','width': '50%', 'display': 'inline-block', 'height': '400px'}),
        dash_table.DataTable(
            id='ma-table',
            columns=[{"name": i, "id": i} for i in df_ma.columns],
            page_size=10,
            style_table={'background-color': '#acecf2','overflowX': 'auto', 'maxHeight': '400px', 'overflowY': 'auto'},
            style_cell={'textAlign': 'left', 'font-size': '8px', 'whiteSpace': 'normal', 'background-color': '#E2EDFF','height': 'auto'},
            style_data_conditional=[{'if': {'column_id': c}, 'width': '100px'} for c in df_ma.columns],
            page_action='native',
            style_as_list_view=True,
        )
    ], style={'display': 'flex'}),

    html.Div([
        html.Label("決算・M＆A情報を選択してください:", style={'font-weight': 'bold'}),
        dcc.Dropdown(
            id='assa-pdf-selector',
            options=[{'label': name, 'value': name} for name in pdf_files.keys()],
            value="2024ーQ2決算情報",  # デフォルト値
            style={'width': '50%'}
        )
    ], style={'margin-bottom': '20px'}),

    # PDFファイル名を表示
    html.Div(id='assa-pdf-filename-display', style={'margin-bottom': '10px', 'font-size': '16px', 'font-weight': 'bold'}),

    # PDF表示用のiframe
    html.Iframe(
        id='assa-pdf-viewer',
        src="",
        style={'width': '80%', 'height': '600px', 'border': 'none','margin': '0 auto', 'display': 'block'}
    ),
    
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
])

# コールバックを定義
@app.callback(
    [Output('assa-sales-ytd-chart', 'figure'),
     Output('assa-ebit-ytd-chart', 'figure'),
     Output('assa-ebit-percent-ytd-chart', 'figure')],
    [Input('assa-currency-type-ytd', 'value'),
     Input('assa-year-checklist', 'value')]
)
def update_ytd_figures(currency_type_ytd,selected_year):
    if not selected_year or not all(year in df_sales_ytd.columns for year in selected_year):
        # 無効な選択があれば空のグラフを返す
        return go.Figure(), go.Figure(), go.Figure()
    
    if currency_type_ytd == 'local':
        sales_ytd_rows = ['EMEIA', 'AMERICAS', 'APAC', 'GLOBAL TECHNOLOGIES', 'ENTRANCE SYSTEMS', 'Other']
        ebit_ytd_rows = ['EMEIA', 'AMERICAS', 'APAC', 'GLOBAL TECHNOLOGIES', 'ENTRANCE SYSTEMS', 'Other']
        group_ytd_row = 'Group'
        group_ytd_row_ebit = 'Group EBIT'
    else:
        sales_ytd_rows = ['EMEIA(JPY)', 'AMERICAS(JPY)', 'APAC(JPY)', 'GLOBAL TECHNOLOGIES(JPY)', 'ENTRANCE SYSTEMS(JPY)', 'Other(JPY)']
        ebit_ytd_rows = ['EMEIA(JPY)', 'AMERICAS(JPY)', 'APAC(JPY)', 'GLOBAL TECHNOLOGIES(JPY)', 'ENTRANCE SYSTEMS(JPY)', 'Other(JPY)']
        group_ytd_row = 'Group(JPY)'
        group_ytd_row_ebit = 'Group EBIT(JPY)'
        
    # Sales YTDグラフ
    sales_ytd_fig = go.Figure()
    sales_ytd_selected = df_sales_ytd.loc[sales_ytd_rows, selected_year]
    for index, row in sales_ytd_selected.iterrows():
        sales_ytd_fig.add_trace(go.Bar(
            x=selected_year,
            y=row,
            name=index,
            offsetgroup=index
        ))
    group_ytd_data = df_sales_ytd.loc[group_ytd_row, selected_year]
    sales_ytd_fig.add_trace(go.Scatter(
        x=selected_year,
        y=group_ytd_data,
        mode='lines+markers+text',
        name='Group',
        line=dict(color='rgba(0, 0, 0, 0.2)'),
        text=group_ytd_data,
        textposition='top center',
        textfont=dict(size=7),
        showlegend=True
    ))
    sales_ytd_fig.update_layout(
        barmode='relative',
        title='Sales-YTD',
        legend=dict(yanchor="top", y=-0.3, xanchor="center", x=0.5)  
    )


    # EBIT YTDグラフ
    ebit_ytd_fig = go.Figure()
    ebit_ytd_selected = df_ebit_ytd.loc[ebit_ytd_rows, selected_year]
    for index, row in ebit_ytd_selected.iterrows():
        ebit_ytd_fig.add_trace(go.Bar(
            x=selected_year,
            y=row,
            name=index,
            offsetgroup=index
        ))
    group_ytd_data_ebit = df_ebit_ytd.loc[group_ytd_row_ebit, selected_year]
    ebit_ytd_fig.add_trace(go.Scatter(
        x=selected_year,
        y=group_ytd_data_ebit,
        mode='lines+markers+text',
        name='Group EBIT',
        line=dict(color='rgba(0, 0, 0, 0.2)'),
        text=group_ytd_data_ebit,
        textposition='top center',
        textfont=dict(size=7),
        showlegend=True
    ))
    ebit_ytd_fig.update_layout(
        barmode='relative',
        title='EBIT-YTD',
        legend=dict(yanchor="top", y=-0.3, xanchor="center", x=0.5)
    )


    # EBIT% YTDグラフ
    ebit_ytd_percent_fig = go.Figure()
    ebit_ytd_percent_rows = ['EMEIA%', 'AMERICAS%', 'APAC%', 'GLOBAL TECHNOLOGIES%', 'ENTRANCE SYSTEMS%', 'Group EBIT %']
    ebit_ytd_percent_selected = df_ebit_ytd.loc[ebit_ytd_percent_rows, selected_year]
    for index, row in ebit_ytd_percent_selected.iterrows():
        ebit_ytd_percent_fig.add_trace(go.Scatter(
            x=selected_year,
            y=row,
            mode='lines+markers',
            name=index
        ))
    ebit_ytd_percent_fig.update_layout(
        title='EBIT YTD%',
        legend=dict(yanchor="top", y=-0.3, xanchor="center", x=0.5)
    )
    ebit_ytd_percent_fig.update_yaxes(tickformat=".1%")

    return sales_ytd_fig, ebit_ytd_fig, ebit_ytd_percent_fig

@app.callback(
    [Output('assa-sales-bar-chart', 'figure'),
     Output('assa-ebit-bar-chart', 'figure'),
     Output('assa-ebit-percent-line-chart', 'figure')],
    [Input('assa-currency-type', 'value'),
     Input('assa-quarter-checklist', 'value')]
)
def update_figures(currency_type, selected_quarters):
    # 通貨タイプに基づいてデータ行を選択
    if currency_type == 'local':
        sales_rows = ['EMEIA', 'AMERICAS', 'APAC', 'GLOBAL TECHNOLOGIES', 'ENTRANCE SYSTEMS', 'Other']
        ebit_rows = ['EMEIA', 'AMERICAS', 'APAC', 'GLOBAL TECHNOLOGIES', 'ENTRANCE SYSTEMS', 'Other']
        group_row = 'Group'
        group_row_ebit = 'Group EBIT'
    else:
        sales_rows = ['EMEIA(JPY)', 'AMERICAS(JPY)', 'APAC(JPY)', 'GLOBAL TECHNOLOGIES(JPY)', 'ENTRANCE SYSTEMS(JPY)', 'Other(JPY)']
        ebit_rows = ['EMEIA(JPY)', 'AMERICAS(JPY)', 'APAC(JPY)', 'GLOBAL TECHNOLOGIES(JPY)', 'ENTRANCE SYSTEMS(JPY)', 'Other(JPY)']
        group_row = 'Group(JPY)'
        group_row_ebit = 'Group EBIT(JPY)'

    # Salesグラフ
    sales_fig = go.Figure()
    sales_selected = df_sales.loc[sales_rows, selected_quarters]
    for index, row in sales_selected.iterrows():
        sales_fig.add_trace(go.Bar(
            x=selected_quarters,
            y=row,
            name=index,
            offsetgroup=index
        ))
    group_data = df_sales.loc[group_row, selected_quarters]
    sales_fig.add_trace(go.Scatter(
        x=selected_quarters,
        y=group_data,
        mode='lines+markers+text',
        name='Group',
        line=dict(color='rgba(0, 0, 0, 0.2)'),
        text=group_data,
        textposition='top center',
        textfont=dict(size=7),
        showlegend=True
    ))
    sales_fig.update_layout(
        barmode='relative',
        title='Sales',
        legend=dict(yanchor="top", y=-0.5, xanchor="center", x=0.5)
    )

    # EBITグラフ
    ebit_fig = go.Figure()
    ebit_selected = df_ebit.loc[ebit_rows, selected_quarters]
    for index, row in ebit_selected.iterrows():
        ebit_fig.add_trace(go.Bar(
            x=selected_quarters,
            y=row,
            name=index,
            offsetgroup=index
        ))
    group_data_ebit = df_ebit.loc[group_row_ebit, selected_quarters]
    ebit_fig.add_trace(go.Scatter(
        x=selected_quarters,
        y=group_data_ebit,
        mode='lines+markers+text',
        name='Group EBIT',
        line=dict(color='rgba(0, 0, 0, 0.2)'),
        text=group_data_ebit,
        textposition='top center',
        textfont=dict(size=7),
        showlegend=True
    ))
    ebit_fig.update_layout(
        barmode='relative',
        title='EBIT',
        legend=dict(yanchor="top", y=-0.5, xanchor="center", x=0.5)
    )

    # EBIT%の折れ線グラフ
    ebit_percent_fig = go.Figure()
    ebit_percent_rows = ['EMEIA%', 'AMERICAS%', 'APAC%', 'GLOBAL TECHNOLOGIES%', 'ENTRANCE SYSTEMS%', 'Group EBIT %']
    ebit_percent_selected = df_ebit.loc[ebit_percent_rows, selected_quarters]
    for index, row in ebit_percent_selected.iterrows():
        ebit_percent_fig.add_trace(go.Scatter(
            x=selected_quarters,
            y=row,
            mode='lines+markers',
            name=index
        ))
    ebit_percent_fig.update_layout(
        title='EBIT%',
        legend=dict(yanchor="top", y=-0.5, xanchor="center", x=0.5)
    )
    ebit_percent_fig.update_yaxes(tickformat=".1%")

    return sales_fig, ebit_fig, ebit_percent_fig

# コールバックを定義して表を更新
@app.callback(
    Output('assa-ma-table', 'data'),
    [Input('assa-year-filter', 'value'),
     Input('assa-sbu-filter', 'value')]
)
def update_ma_table(selected_years, selected_sbus):
    # データをフィルタリング
    filtered_ma_df = df_ma.copy()
    if selected_years:
        filtered_ma_df = filtered_ma_df[filtered_ma_df['Year'].isin(selected_years)]
    if selected_sbus:
        filtered_ma_df = filtered_ma_df[filtered_ma_df['SBU'].isin(selected_sbus)]
    # フィルタリングされたデータを返す
    return filtered_ma_df.to_dict('records')

@app.callback(
    Output('assa-choropleth-map', 'figure'),
    [Input('assa-year-filter', 'value'),
     Input('assa-sbu-filter', 'value')]
)
def update_map(selected_years, selected_sbus):
    # データをフィルタリング
    filtered_df = df_ma.copy()
    if selected_years:
        filtered_df = filtered_df[filtered_df['Year'].isin(selected_years)]
    if selected_sbus:
        filtered_df = filtered_df[filtered_df['SBU'].isin(selected_sbus)]
    # データが空でないか確認
    if filtered_df.empty:
        return go.Figure()
    # Choropleth地図を作成
    fig = px.choropleth(
        filtered_df,
        locations="Country",
        locationmode="country names",
        color="Sales PY",
        hover_name="Company",
        hover_data={
            "Date": True,
            "Industry": True,
            "Category": True,
            "Strategy": True,
            "Sales PY": True,
            "HC": True,
            "Country": True,
            "City": True,
        },
        title="M&A Map",
        projection="natural earth"
    )
    # 地図のレイアウトを更新
    fig.update_geos(
        showcoastlines=True,
        showland=True,
    )
    fig.update_layout(
        geo=dict(bgcolor='white'),
        margin={"r": 0, "t": 0, "l": 0, "b": 0},
        coloraxis_showscale=False
    )
    return fig

# Sales Growthグラフのコールバック
@app.callback(
    Output('assa-sales-growth-line-chart', 'figure'),
    [Input('assa-sales-growth-type', 'value')]
)
def update_sales_growth(selected_type):
    # 選択されたタイプに基づいてデータをフィルタリング
    if selected_type == 'group':
        sales_growth_rows = ['Prise(Group)', 'Volume(Group)', 'FX(Group)', 'Acq/Div(Group)']
    else:
        sales_growth_rows = ['Organic sales(Entrance)', 'FX(Entrance)', 'Acq/Div(Entrance)']
    sales_growth_data = df_q_bridge.loc[sales_growth_rows]
    # 折れ線グラフを作成
    sales_growth_fig = go.Figure()
    for index, row in sales_growth_data.iterrows():
        sales_growth_fig.add_trace(go.Scatter(
            x=sales_growth_data.columns,
            y=row,
            mode='lines+markers',
            name=index
        ))
    sales_growth_fig.update_layout(
        title='Sales Growth',
        legend=dict(yanchor="top", y=-0.3, xanchor="center", x=0.5)
    )
    sales_growth_fig.update_yaxes(tickformat=".0%")
    return sales_growth_fig

@app.callback(
    [Output('assa-pdf-viewer', 'src'), Output('pdf-filename-display', 'children')],
    [Input('assa-pdf-selector', 'value')]
)
def update_pdf_viewer(selected_pdf_name):
    # 選択されたPDFのパスを取得
    selected_pdf_path = pdf_files.get(selected_pdf_name)

    # PDFをエンコード
    if selected_pdf_path:
        encoded_pdf = encode_pdf(selected_pdf_path)
        if encoded_pdf:
            # PDFのファイル名を表示
            return encoded_pdf, f"表示中のPDF: {selected_pdf_name}"
        else:
            return "", "エラー: PDFファイルが見つかりません。"
    else:
        return "", "エラー: 選択されたPDFオプションが無効です。"

# アプリを実行
if __name__ == '__main__':
    app.layout = create_layout(app)  # レイアウトを設定
    app.run_server(debug=True)