import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import plotly.express as px
import dash_bootstrap_components as dbc
from dash import dash_table  # dash_tableをインポート

# 各シートのデータを読み込む
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Manusa\Manusa.xlsx"
sales_sheet = "Sales"
info_sheet = "Information"
news_sheet = "News"
agents_sheet = "代理店"

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)

df_news = pd.read_excel(file_path, sheet_name=news_sheet, header=0)
df_news.fillna('', inplace=True)

sales_data = pd.read_excel(file_path, sheet_name=sales_sheet)

# "代理店"シートのデータを読み込む
df_agents = pd.read_excel(file_path, sheet_name=agents_sheet, header=0)
df_agents.fillna('', inplace=True)

df_agents[['緯度', '経度']] = df_agents['経緯度'].str.split(',', expand=True)
df_agents['緯度'] = df_agents['緯度'].astype(float)
df_agents['経度'] = df_agents['経度'].astype(float)

# 必要な列の定義
columns = ["2020", "2021", "2022"]

# 現地通貨データの抽出
group_local = sales_data[sales_data["Sales"] == "Group"][columns].values[0]
adjusted_ebit_local = sales_data[sales_data["Sales"] == "Group EBIT"][columns].values[0]
group_ebit_percentage = sales_data[sales_data["Sales"] == "Group EBIT %"][columns].values[0]

# 日本円換算データの抽出
group_jpy = sales_data[sales_data["Sales"] == "Group(JPY)"][columns].values[0]
adjusted_ebit_jpy = sales_data[sales_data["Sales"] == "Group EBIT(JPY)"][columns].values[0]

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

# 地図を生成する関数
def generate_map(dataframe):
    # Plotly Expressで地図を作成
    fig = px.scatter_mapbox(
        dataframe,
        lat='緯度',
        lon='経度',
        text='Company',
        hover_data={'Company': True, 'Link': True, 'City': True, 'Nation': True},
        zoom=2,
        height=600
    )

    # hovertemplateの修正
    fig.update_traces(
        hovertemplate=(
            '<b>%{text}</b><br>' +
            'City: %{customdata[2]}<br>' +
            'Nation: %{customdata[3]}<br>' +
            'Link: <a href="%{customdata[1]}" target="_blank" style="color:yellow; font-weight:bold;">%{customdata[1]}</a>'
        ),
    )

    # レイアウトの更新
    fig.update_layout(
        mapbox_style="carto-positron",  # 地図スタイルを指定
        margin={"r": 0, "t": 0, "l": 0, "b": 0},
        title='代理店マップ',
        title_x=0.5  # タイトルを中央揃え
    )

    return fig

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

image_url = "https://cdn1.npcdn.net/images/156967039337b86d8e72ebb5a29e5b3ddb47815c7c.jpg?md5id=1ecb6468735bde11b1d125b787158088&amp;amp;new_width=1920&amp;amp;new_height=1920&amp;amp;w=-62170009200"

# 代理店データを表示するテーブルを生成する関数
def generate_agents_table(dataframe):
    return dash_table.DataTable(
        data=dataframe.to_dict('records'),
        columns=[{"name": i, "id": i} for i in dataframe.columns],
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'left',
            'padding': '3px',
             "fontSize": "8px",  # 设置成较小的字体如10px，可根据需要调整
        },
        page_size=10,  # ページサイズ
        style_header={
            'backgroundColor': 'lightblue',
            'fontWeight': 'bold'
        }
    )

# アプリケーションのレイアウト
def create_layout(app):
    return html.Div([
    html.Div([
        html.Div(generate_info_content(df_info), style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-bottom': '0px', 'flex': '3'}),
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
            {'label': '現地通貨(MEUR)', 'value': 'local'},
            {'label': '日本円換算(億円)', 'value': 'jpy'}
        ],
        value='local',  # デフォルト値
        labelStyle={'display': 'inline-block', 'margin-right': '10px'},
        style={'margin-bottom': '20px'}
    ),

    dcc.Graph(id='sales-chart', style={'height': '400px'}),

    html.Div(html.H2("代理店マップ", style={
        'text-align': 'center',
        'font-size': '24px',
        'font-weight': 'bold',
        'margin-bottom': '20px',
        'color': 'black',
        'background-color': '#ADD8E6',
        'padding': '10px',
        'border-radius': '5px'
    })),

    dcc.Graph(id='agents-map', figure=generate_map(df_agents)),  # 地図

    # 代理店データを表示するテーブルを追加
    html.Div(generate_agents_table(df_agents), style={'margin': '20px 0'}),

    html.Div(id="dummy-output", style={"display": "none"}),
    
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
])

@app.callback(
    Output("dummy-output", "children"),
    [Input("agents-map", "clickData")]
)
def open_link(click_data):
    if click_data is None:
        return None
    # クリックされたデータからリンクを取得
    link = click_data['points'][0]['customdata']
    return html.Script(f'window.open("{link}", "_blank");')

@app.callback(
    Output("sales-chart", "figure"),
    [Input("currency-type-ytd", "value")]
)
def update_chart(currency):
    if currency == "local":
        group_data = group_local
        ebit = adjusted_ebit_local
        ebit_percentage = group_ebit_percentage  # Group EBIT % のデータを取得
        yaxis_title = "Value (MEUR)"
    elif currency == "jpy":
        group_data = group_jpy
        ebit = adjusted_ebit_jpy
        ebit_percentage = group_ebit_percentage  # Group EBIT % のデータを取得
        yaxis_title = "Value (億円)"

    fig = go.Figure()
    half_years = [f"{col}" for col in columns]

    # Salesの棒グラフ
    fig.add_trace(go.Bar(
        x=half_years,
        y=group_data,
        name="Sales",
        marker_color="pink"
    ))

    # EBITの棒グラフ
    fig.add_trace(go.Bar(
        x=half_years,
        y=ebit,
        name="EBIT",
        marker_color="green"
    ))

    # Group EBIT % の折れ線グラフ
    fig.add_trace(go.Scatter(
        x=half_years,
        y=ebit_percentage,
        name="EBIT %",
        mode="lines+markers",
        marker=dict(color="blue", size=10),
        line=dict(width=2),
        yaxis="y2"  # ここで第二Y軸を指定
    ))

    # グラフのレイアウト更新
    fig.update_layout(
        barmode="group",  # ここでグループモードを指定
        title="Sales Data Visualization",
        xaxis_title="Period",
        yaxis_title=yaxis_title,
        yaxis2=dict(
            title="EBIT %",
            overlaying='y',
            side='right',
            showgrid=False,
            zeroline=False,
            titlefont=dict(color="blue"),
            tickfont=dict(color="blue")
        ),
        legend_title="Legend",
        template="plotly_white"
    )

    return fig

if __name__ == '__main__':
    app.run_server(debug=True)