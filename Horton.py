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

# PDFファイルのパス
pdf_files = {
    "2024ーQ1決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Horton\三和HD業績2024-Q1.pdf",
    "2024ー1H決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Horton\三和HD業績2024-1H.pdf",
    "2024ーQ3決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Horton\三和HD業績2024-Q3.pdf",
    "2024歩行者決算": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Horton\三和ー歩行者業績まとめ.pdf",
    "M&A情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Horton\三和M&A.pdf",
    "生産拠点": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Horton\三和生産拠点.pdf"
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
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Horton\Horton.xlsx"
sales_sheet = "Sales"
info_sheet = "Information"
news_sheet = "News"

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)
df_news = pd.read_excel(file_path, sheet_name=news_sheet, header=0)
df_news.fillna('', inplace=True)

# Salesデータの読み込み
sales_data = pd.read_excel(file_path, sheet_name=sales_sheet)

# 必要な列の定義
columns = ["2021", "2022", "2023"]

# 現地通貨データの抽出
group_local = sales_data[sales_data["Sales"] == "Group"][columns].values[0]

# 日本円換算データの抽出
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

image_url = "https://www.hortonaccess.com/assets/images/content/hd-pas_brand-logos_2x.png"

# アプリケーションのレイアウト
def create_layout(app):
    return html.Div([
    html.Div([
        html.Div(generate_info_content(df_info), style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-bottom': '0px', 'flex': '3'}),
        html.Div([html.Img(src=image_url, style={'width': '100%', 'height': '100%', 'margin-bottom': '0px'})], style={'flex': '1', 'padding': '1px'})
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
    html.Div([
        html.Label("決算・M＆A情報を選択してください:", style={'font-weight': 'bold'}),
        dcc.Dropdown(
            id='pdf-selector',
            options=[{'label': name, 'value': name} for name in pdf_files.keys()],
            value="2024ーQ1決算情報",  # デフォルト値
            style={'width': '50%'}
        )
    ], style={'margin-bottom': '20px'}),

    # PDFファイル名を表示
    html.Div(id='pdf-filename-display', style={'margin-bottom': '10px', 'font-size': '16px', 'font-weight': 'bold'}),

    # PDF表示用のiframe
    html.Iframe(
        id='pdf-viewer',
        src="",
        style={'width': '80%', 'height': '600px', 'border': 'none','margin': '0 auto', 'display': 'block'}
    ),
    
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
])
@app.callback(
    [Output('pdf-viewer', 'src'), Output('pdf-filename-display', 'children')],
    [Input('pdf-selector', 'value')]
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
    
@app.callback(
    Output("sales-chart", "figure"),
    [Input("currency-type-ytd", "value")]
)
def update_chart(currency):
    if currency == "local":
        group_data = group_local
        yaxis_title = "Value (MUSD)"
    elif currency == "jpy":
        group_data = group_jpy
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

    
    # グラフのレイアウト更新
    fig.update_layout(
        barmode="group",  # ここでグループモードを指定
        title="Sales Data Visualization",
        xaxis_title="Period",
        yaxis_title=yaxis_title,
        legend_title="Legend",
        template="plotly_white"
    )

    return fig

if __name__ == '__main__':
    app.run_server(debug=True)