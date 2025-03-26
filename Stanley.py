import base64
import pandas as pd
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import dash_bootstrap_components as dbc

# PDFファイルのパス
pdf_files = {
    "AllegionーQ2決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Stanley\AllegionQ2決算.pdf",
    "AllegionーQ3決算情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Stanley\AllegionQ3決算.pdf",
    "M&A情報": r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Stanley\Allegion M&A.pdf"
}

# PDFをBase64エンコードする関数
def encode_pdf(file_path):
    try:
        with open(file_path, 'rb') as pdf_file:
            encoded_pdf = base64.b64encode(pdf_file.read()).decode('utf-8')
        return f"data:application/pdf;base64,{encoded_pdf}"
    except FileNotFoundError:
        print(f"PDFファイルが見つかりません: {file_path}")
        return None
    except Exception as e:
        print(f"PDFエンコード中にエラーが発生しました: {e}")
        return None

# 各シートのデータを読み込む
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Stanley\Stanley.xlsx"
sales_sheet = "Sales"
info_sheet = "Information"
news_sheet = "News"

try:
    df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
    df_info.fillna('', inplace=True)
    df_news = pd.read_excel(file_path, sheet_name=news_sheet, header=0)
    df_news.fillna('', inplace=True)
    sales_data = pd.read_excel(file_path, sheet_name=sales_sheet)
except FileNotFoundError:
    print(f"Excelファイルが見つかりません: {file_path}")
    df_info, df_news, sales_data = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# 必要な列の定義
columns = ["2020", "2021", "2022", "2023"]

# データ抽出
try:
    group_local = sales_data[sales_data["Sales"] == "Group"][columns].values[0]
    adjusted_ebit_local = sales_data[sales_data["Sales"] == "Group EBIT"][columns].values[0]
    group_ebit_percentage = sales_data[sales_data["Sales"] == "Group EBIT %"][columns].values[0]
    group_jpy = sales_data[sales_data["Sales"] == "Group(JPY)"][columns].values[0]
    adjusted_ebit_jpy = sales_data[sales_data["Sales"] == "Group EBIT(JPY)"][columns].values[0]
except IndexError:
    print("データ抽出中にエラーが発生しました。Excelファイルの内容を確認してください。")
    group_local, adjusted_ebit_local, group_ebit_percentage = [], [], []
    group_jpy, adjusted_ebit_jpy = [], []

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

def create_layout(app):
    layout = html.Div([
        html.Div([
            html.Div(generate_info_content(df_info), style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-bottom': '0px', 'flex': '3'}),
            html.Div([html.Img(src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSC96NpW-G6Tjlp3U-Xo9OFEuKOjZXU-vd4eg&s", style={'width': '80%', 'height': '80%', 'margin-bottom': '0px'})], style={'flex': '1', 'padding': '1px'})
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
            value='local',
            labelStyle={'display': 'inline-block', 'margin-right': '10px'},
            style={'margin-bottom': '20px'}
        ),
        dcc.Graph(id='sales-chart', style={'height': '400px'}),
        html.Div([
            html.Label("決算・M＆A情報を選択してください:", style={'font-weight': 'bold'}),
            dcc.Dropdown(
                id='pdf-selector',
                options=[{'label': name, 'value': name} for name in pdf_files.keys()],
                value="AllegionーQ2決算情報",
                style={'width': '50%'}
            )
        ], style={'margin-bottom': '20px'}),
        html.Div(id='pdf-filename-display', style={'margin-bottom': '10px', 'font-size': '16px', 'font-weight': 'bold'}),
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
        Output("sales-chart", "figure"),
        [Input("currency-type-ytd", "value")]
    )
    def update_chart(currency):
        if currency == "local":
            group_data = group_local
            ebit = adjusted_ebit_local
            ebit_percentage = group_ebit_percentage
            yaxis_title = "Value (MEUR)"
        elif currency == "jpy":
            group_data = group_jpy
            ebit = adjusted_ebit_jpy
            ebit_percentage = group_ebit_percentage
            yaxis_title = "Value (億円)"
        if not group_data or not ebit or not ebit_percentage:
            print("データが空です。グラフを生成できません。")
            return go.Figure()
        fig = go.Figure()
        half_years = [f"{col}" for col in columns]
        fig.add_trace(go.Bar(x=half_years, y=group_data, name="Sales", marker_color="pink"))
        fig.add_trace(go.Bar(x=half_years, y=ebit, name="EBIT", marker_color="green"))
        fig.add_trace(go.Scatter(
            x=half_years, y=ebit_percentage, name="EBIT %",
            mode="lines+markers", marker=dict(color="blue", size=10),
            line=dict(width=2), yaxis="y2"
        ))
        fig.update_layout(
            barmode="group",
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

    @app.callback(
        [Output('pdf-viewer', 'src'), Output('pdf-filename-display', 'children')],
        [Input('pdf-selector', 'value')]
    )
    def update_pdf_viewer(selected_pdf_name):
        selected_pdf_path = pdf_files.get(selected_pdf_name)
        if selected_pdf_path:
            encoded_pdf = encode_pdf(selected_pdf_path)
            if encoded_pdf:
                return encoded_pdf, f"表示中のPDF: {selected_pdf_name}"
            else:
                return "", "エラー: PDFファイルが見つかりません。"
        else:
            return "", "エラー: 選択されたPDFオプションが無効です。"

    return layout