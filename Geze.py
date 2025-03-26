import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Excelファイルのパス
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\GEZE\Geze.xlsx"

# ----------------------------------------
# 第一部：Choropleth Map と データテーブル (直販シート)
# ----------------------------------------
df_map = pd.read_excel(file_path, sheet_name="直販", engine="openpyxl")

# Branch列に基づいた色分けマッピング
branch_colors = {
    "Germany": "blue",
    "Spain": "red",
    "France": "green",
    "Italy": "purple",
    # 他のBranchも必要に応じて追加
}

# ----------------------------------------
# 第二部：Sales, Information, News シート
# ----------------------------------------
info_sheet = "Information"
news_sheet = "News"
sales_sheet = "Sales"

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)
df_news = pd.read_excel(file_path, sheet_name=news_sheet, header=0)
df_news.fillna('', inplace=True)

sales_data = pd.read_excel(file_path, sheet_name=sales_sheet)
columns = ["2021", "2022", "2023"]

# 現地通貨データの抽出
group_local = sales_data[sales_data["Sales"] == "Group"][columns].values[0]
adjusted_ebit_local = sales_data[sales_data["Sales"] == "Group EBIT"][columns].values[0]
group_ebit_percentage = sales_data[sales_data["Sales"] == "Group EBIT %"][columns].values[0]

# 日本円換算データの抽出
group_jpy = sales_data[sales_data["Sales"] == "Group(JPY)"][columns].values[0]
adjusted_ebit_jpy = sales_data[sales_data["Sales"] == "Group EBIT(JPY)"][columns].values[0]

# Informationコンテンツ生成用の関数
def generate_info_content(dataframe):
    content = []
    for _, row in dataframe.iterrows():
        content.append(html.P([
            html.Strong(f"{row['テーマ']}: "),
            html.Span(row['内容'])
        ], style={'font-size': '10px', 'margin': '0px 0'}))
    return content

# Newsコンテンツ生成用の関数
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

# Dashアプリケーションの作成（Bootstrap用の外部スタイルシートを利用）
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# GEZEロゴ画像
image_url = "https://upload.wikimedia.org/wikipedia/commons/a/a3/GEZE_logo.svg"

# アプリケーションのレイアウト
def create_layout(app):
    return html.Div([
    

    # -----------------------------
    # 第二部：Information, News, Sales Chart
    # -----------------------------
    html.Div([
        html.Div([
            html.Div(generate_info_content(df_info),
                     style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-bottom': '0px', 'flex': '3'}),
            html.Div([html.Img(src=image_url, style={'width': '80%', 'height': '80%', 'margin-bottom': '0px'})],
                     style={'flex': '1', 'padding': '1px'})
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

        html.Div(generate_news_content(df_news),
                 style={'background-color': '#f9f9f9', 'padding': '3px', 'margin': '0px'}),

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

        dcc.Graph(id='sales-chart', style={'height': '400px'})
    ], style={"margin": "20px"}),
    # -----------------------------
    # 第一部：Choropleth Map with Search と データテーブル
    # -----------------------------
    html.Div([
        html.H1("直販カーバ地域", style={"textAlign": "center"}),
        html.Div([
            dcc.Input(
                id="search-bar",
                type="text",
                placeholder="Enter search keyword...",
                style={"width": "50%", "padding": "10px"}
            ),
            html.Button("Search", id="search-button", n_clicks=0, style={"marginLeft": "10px"})
        ], style={"textAlign": "center", "marginBottom": "20px"}),
        dcc.Graph(
            id="choropleth-map",
            style={"height": "60vh"}
        ),
        dash_table.DataTable(
            id="data-table",
            columns=[{"name": col, "id": col} for col in df_map.columns],
            data=df_map.to_dict("records"),
            page_size=8,
            style_table={"height": "40vh", "overflowY": "auto"},
            style_cell={"textAlign": "left",
                        "fontSize": "10px",  # 设置成较小的字体如10px，可根据需要调整
                        "padding": "0px"},
        ),
    ], style={"margin": "10px"}),
    
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        }),

    # 区切り
    html.Hr(),
])

# ----------------------------------------
# コールバック関数（Choropleth Mapとデータテーブル更新）
# ----------------------------------------
@app.callback(
    [Output("choropleth-map", "figure"),
     Output("data-table", "data")],
    [Input("search-button", "n_clicks")],
    [State("search-bar", "value")]
)
def update_content(n_clicks, search_value):
    if not search_value:
        filtered_df = df_map
    else:
        # 全列に対して検索キーワードが含まれているかチェック
        filtered_df = df_map[df_map.apply(lambda row: row.astype(str).str.contains(search_value, case=False).any(), axis=1)]

    fig = px.choropleth(
        filtered_df,
        locations="Country",  # 国名を指定
        locationmode="country names",
        color="Branch",       # Branch列による色分け
        hover_name="Country",
        color_discrete_map=branch_colors
    )
    fig.update_geos(projection_type="natural earth")
    fig.update_layout(margin={"r": 0, "t": 50, "l": 0, "b": 0})

    return fig, filtered_df.to_dict("records")

# ----------------------------------------
# コールバック関数（Sales Chart更新）
# ----------------------------------------
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

    # Group EBIT % の折れ線グラフ（第二Y軸）
    fig.add_trace(go.Scatter(
        x=half_years,
        y=ebit_percentage,
        name="EBIT %",
        mode="lines+markers",
        marker=dict(color="blue", size=10),
        line=dict(width=2),
        yaxis="y2"
    ))

    fig.update_layout(
        barmode="group",
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
            tickformat=".1%",
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

if __name__ == '__main__':
    app.run_server(debug=True)