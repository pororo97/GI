import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import dash_bootstrap_components as dbc

# 各シートのデータを読み込む
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\GU\GU.xlsx"
sales_sheet = "Sales"
info_sheet = "Information"
time_sheet = "time"  # 追加: timeシートの名前

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)

df_time = pd.read_excel(file_path, sheet_name=time_sheet, header=0)  # 追加: timeシートのデータを読み込み
df_time.fillna('', inplace=True)  # NaNを空文字に置き換え

# Salesデータの読み込み
sales_data = pd.read_excel(file_path, sheet_name=sales_sheet)

# 必要な列の定義
columns = ["2015", "2016", "2017"]

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
        ], style={'font-size': '10px', 'margin': '1px 0'}))
    return content

def generate_time_content(dataframe):
    content = []
    for _, row in dataframe.iterrows():
        content.append(html.P([
            html.Strong(f"{row['日時']}: "),  # テーマ列を太字に
            html.Span(row['内容'])
        ], style={'font-size': '10px', 'margin': '1px 0'}))
    return content

# Dashアプリケーションの作成
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

image_url = "https://www.gu-bks.de/fileadmin/user_upload/GU-BKS/Unternehmen/gu_netzwerk.jpg"
new_image_url = "https://www.gu-bks.de/fileadmin/_processed_/b/a/csm_gu_firmen_weiss2_e846522d9b.jpg"

# アプリケーションのレイアウト
def create_layout(app):
    return html.Div([
    html.Div([
        html.Div(generate_info_content(df_info), style={'background-color': '#CDE1FF', 'font-size': '12px', 'padding': '3px', 'margin-bottom': '0px', 'flex': '3'}),
        html.Div([html.Img(src=image_url, style={'width': '100%', 'height': '100%', 'margin-bottom': '2px'})], style={'flex': '1', 'padding': '0px'})
    ], style={'display': 'flex', 'justify-content': 'space-between'}),

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

    # 新しい画像のタイトルと画像を追加
    html.Div([
        html.H4("代理店一覧", style={'text-align': 'center', 'margin-top': '20px'}),  # タイトルの追加
        html.Img(src=new_image_url, style={'width': '90%', 'height': '60%', 'margin-top': '10px'})
    ], style={'text-align': 'center'}),

    # 会社組織変更のセクションを追加
    html.Div([
        html.H4("会社組織変更", style={'text-align': 'center', 'margin-top': '20px'}),
        html.Div(generate_time_content(df_time), style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-top': '10px'})
    ], style={'text-align': 'lift'}),
    
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
    print ('fig')

    # グラフのレイアウト更新
    fig.update_layout(
    barmode="group",  # ここでグループモードを指定
    title="業績まとめ",
    xaxis_title="Period",
    yaxis_title=yaxis_title,
        yaxis2=dict(
        title="EBIT %",
        overlaying='y',
        side='right',
        showgrid=False,
        zeroline=False,
        titlefont=dict(color="blue"),
        tickfont=dict(color="blue"),
        tickformat=".1%",  # ここでパーセント表示を指定
        title_standoff=10
    ),
       legend_title="Legend",
       template="plotly_white",
       legend=dict(
        x=1.1,
        y=1.0,
        xanchor="left",
        yanchor="top",
        font=dict(size=10)
    ),
)

    return fig

# サーバーを実行
if __name__ == '__main__':
    app.run_server(debug=True)