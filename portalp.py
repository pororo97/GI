import pandas as pd
import dash
from dash import dcc, html
import plotly.graph_objects as go
import dash_bootstrap_components as dbc
import networkx as nx

# 各シートのデータを読み込む
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\portalp\portalp.xlsx"
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

image_url = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRahgmLrHL7bmVN6fZTRYkOtwBUYiBqqtN7qA&s"
new_image_url = "https://www.portalp.com/wp-content/uploads/2024/10/portalp-mappemonde-distibuteurs-filiales-full-en-2.jpg"


# アプリケーションのレイアウト
def create_layout(app):
    return html.Div([
        # 第一部分：信息内容和图片
        html.Div([
            html.Div(generate_info_content(df_info), style={
                'background-color': '#CDE1FF', 
                'font-size': '8px', 
                'padding': '5px', 
                'margin-bottom': '0px', 
                'flex': '3'
            }),
            html.Div([
                html.Img(src=image_url, style={
                    'width': '100%', 
                    'height': '100%', 
                    'margin-bottom': '1px'
                })
            ], style={'flex': '1', 'padding': '2px'})
        ], style={'display': 'flex', 'justify-content': 'space-between'}),

        # 第二部分：新的图片和标题
        html.Div([
            html.H4("代理店一覧", style={'text-align': 'center', 'margin-top': '20px'}),  # 标题
            html.Img(src=new_image_url, style={
                'width': '90%', 
                'height': '100%', 
                'margin-top': '10px'
            })
        ], style={'text-align': 'center'}),

        # 返回目录页的链接
        html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
    ])

# サーバーを実行
if __name__ == '__main__':
    app.run_server(debug=True)