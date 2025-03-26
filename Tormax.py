import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import dash_bootstrap_components as dbc
import networkx as nx
import webbrowser

# 各シートのデータを読み込む
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\Tormax\Tormax.xlsx"
info_sheet = "Information"

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)

def generate_info_content(dataframe):
    content = []
    for _, row in dataframe.iterrows():
        content.append(html.P([
            html.Strong(f"{row['テーマ']}: "),  # テーマ列を太字に
            html.Span(row['内容'])
        ], style={'font-size': '10px', 'margin': '1px 0'}))
    return content

# Dashアプリケーションの作成
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

image_url = "https://lh3.googleusercontent.com/proxy/LRDsbiE6LaDUFkFlwneUQ57huJggnLeSkE_A7yJzpt01zGrxbIPUALKWYpKs8fRN9crlL-A7YD3X8iRxTkZDZzWPd-dkKWIItxBJWNvRnIcudxLWdw"


# アプリケーションのレイアウト
def create_layout(app):
    return html.Div([
    html.Div([
        html.Div(generate_info_content(df_info), style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-bottom': '0px', 'flex': '3'}),
        html.Div([html.Img(src=image_url, style={'width': '90%', 'height': '60%', 'margin-bottom': '1px'})], style={'flex': '1', 'padding': '2px'})
    ], style={'display': 'flex', 'justify-content': 'space-between'}),
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
])


# サーバーを実行
if __name__ == '__main__':
    app.run_server(debug=True)