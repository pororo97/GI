import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import dash_bootstrap_components as dbc
import networkx as nx
import webbrowser

# 各シートのデータを読み込む
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\FAAC\FAAC.xlsx"
sales_sheet = "Sales"
info_sheet = "Information"

df_info = pd.read_excel(file_path, sheet_name=info_sheet, header=0)
df_info.fillna('', inplace=True)

# Salesデータの読み込み
sales_data = pd.read_excel(file_path, sheet_name=sales_sheet)

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
        ], style={'font-size': '10px', 'margin': '1px 0'}))
    return content

# Dashアプリケーションの作成
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

image_url = "https://faac.blob.core.windows.net/web/1/root/faac-technologies-positive.png"

# Access Controlの関係図データの定義
nodes_access = {
    "Access Control": {"info": "", "link": ""},
    "MAGNETIC": {"info": "最先端の使いやすい入退室管理ソリューションに特化した「メイド・イン・ドイツ」の世界トップブランド", "link": "http://www.magnetic-access.com"},
    "WOLPAC": {"info": "北欧およびスカンジナビア諸国のマーケットリーダーとして、DAABは商業および産業活動のための工業規模のオートメーションサービスを製造しています。", "link": "http://www.wolpac.com.br"},
}

# Parkingの関係図データの定義
nodes_automation = {
    "Automation": {"info": "", "link": ""},
    "FAAC": {"info": "ヨーロッパ市場および非ヨーロッパ諸国、特に住宅用ゲートオートメーション分野におけるリーダーです。FAACテクノロジーズ社の研究開発の中心であり", "link": "http://www.faac.co.uk"},
    "DAAB": {"info": "北欧およびスカンジナビア諸国のマーケットリーダーとして、DAABは商業および産業活動のための工業規模のオートメーションサービスを製造しています。", "link": "http://www.faac.se"},
    "CLEMSA": {"info": "スペイン市場をリードするブランド", "link": "http://www.clemsa.es"},
    "GENIUS": {"info": "電気機械技術に焦点を当て、提案されるソリューションが提供する優れたコストパフォーマンスを特徴とするグループのブランド", "link": "http://geniusg.com/en"},
    "ROSSI": {"info": "ブラジルとラテンアメリカ市場のリーダーであるこのブランドは、ゲート、バリア、ガレージのオートメーション・システム、電子回路基板、遠隔制御装置の設計と施工を専門としている。", "link": "http://www.rossiportoes.com.br"},
    "CENTURION": {"info": "南アフリカとアフリカのトップブランド。ゲート、バリア、安全装置、電子アクセス・コントロール装置のオペレーターを設計・製造している。", "link": "http://www.centsys.co.za"},
    "VIKING": {"info": "北米市場、特にアメリカにおいて、スライディングゲートとスイングゲートのオペレーターを専門とするトップブランド", "link": "http://www.vikingaccess.com"},
    "SIMPLY CONNECT": {"info": "FAACがインストーラー向けに開発したアプリで、インストーラーがいつでもどこからでも顧客のオートメーションシステムと遠隔操作できるように設計されています", "link": ""},
    "COMETA": {"info": "セキュリティ＆セーフティソリューション分野", "link": "http://www.cometaspa.com"},
    "TECHNO FIRE": {"info": "あらゆる種類のテクニカルクロージャーの販売、設置、メンテナンス、修理を専門とするイタリアの企業である。", "link": "https://www.techno-fire.com"}
}

# Access Controlのグラフを作成する関数
def create_access_control_graph():
    G = nx.DiGraph()
    G.add_node("Access Control", layer=0)
    for node in nodes_access:
        if node != "Access Control":
            G.add_node(node, layer=1)
            G.add_edge("Access Control", node)

    pos = nx.multipartite_layout(G, subset_key="layer")
    for node in pos:
        pos[node] = (pos[node][0] * 0.5, pos[node][1] * 0.5)
    edge_x = []
    edge_y = []
    for edge in G.edges():
        x0, y0 = pos[edge[0]]
        x1, y1 = pos[edge[1]]
        edge_x.append(x0)
        edge_x.append(x1)
        edge_x.append(None)
        edge_y.append(y0)
        edge_y.append(y1)
        edge_y.append(None)

    edge_trace = go.Scatter(
        x=edge_x, y=edge_y,
        line=dict(width=2, color='#666'),
        hoverinfo='none',
        mode='lines'
    )

    node_x = []
    node_y = []
    hover_text = []
    for node in G.nodes():
        x, y = pos[node]
        node_x.append(x)
        node_y.append(y)
        info = nodes_access[node]["info"]
        link = nodes_access[node]["link"]
        formatted_info = "<br>".join(info[i:i + 80] for i in range(0, len(info), 80))
        if link:
            hover_text.append(f"{node}: {formatted_info}<br><a href='{link}' target='_blank'>リンクを見る</a>")
        else:
            hover_text.append(f"{node}: {formatted_info}")

    node_trace = go.Scatter(
        x=node_x, y=node_y,
        mode='markers+text',
        text=[f"{node}" for node in G.nodes()],
        textposition="middle right",
        hoverinfo='text',
        marker=dict(
            showscale=False,
            colorscale='YlGnBu',
            size=5,
        ),
    )
    node_trace.text = hover_text

    fig = go.Figure(data=[edge_trace, node_trace],
                    layout=go.Layout(
                        title="Access Control",
                        titlefont_size=4,
                        showlegend=False,
                        hovermode='closest',
                        margin=dict(b=3, l=5, r=0, t=0),
                        xaxis=dict(
                            showgrid=False,
                            zeroline=False,
                            showticklabels=False,
                            range=[-1.1, 5]  # x軸の表示範囲を制限
                        ),
                        yaxis=dict(
                            showgrid=False,
                            zeroline=False,
                            showticklabels=False,
                            range=[-1, 1]  # y軸の表示範囲を制限
                        ),
                        dragmode='pan'
                    )
                    )
    return fig

# Parkingのグラフを作成する関数
def create_automation_graph():
    G = nx.DiGraph()
    G.add_node("Automation", layer=0)
    for node in nodes_automation:
        if node != "Automation":
            G.add_node(node, layer=1)
            G.add_edge("Automation", node)

    pos = nx.multipartite_layout(G, subset_key="layer")
    for node in pos:
        pos[node] = (pos[node][0] * 1.1, pos[node][1] * 0.8)

    edge_x = []
    edge_y = []
    for edge in G.edges():
        x0, y0 = pos[edge[0]]
        x1, y1 = pos[edge[1]]
        edge_x.append(x0)
        edge_x.append(x1)
        edge_x.append(None)
        edge_y.append(y0)
        edge_y.append(y1)
        edge_y.append(None)

    edge_trace = go.Scatter(
        x=edge_x, y=edge_y,
        line=dict(width=2, color='#888'),
        hoverinfo='none',
        mode='lines'
    )

    node_x = []
    node_y = []
    hover_text = []
    for node in G.nodes():
        x, y = pos[node]
        node_x.append(x)
        node_y.append(y)
        info = nodes_automation[node]["info"]
        link = nodes_automation[node]["link"]
        formatted_info = "<br>".join(info[i:i + 100] for i in range(0, len(info), 100))
        if link:
            hover_text.append(f"{node}: {formatted_info}<br><a href='{link}' target='_blank'>リンクを見る</a>")
        else:
            hover_text.append(f"{node}: {formatted_info}")

    node_trace = go.Scatter(
        x=node_x, y=node_y,
        mode='markers+text',
        text=[f"{node}" for node in G.nodes()],
        textposition="middle right",
        hoverinfo='text',
        marker=dict(
            showscale=False,
            colorscale='YlGnBu',
            size=4,
        ),
    )
    node_trace.text = hover_text

    fig = go.Figure(data=[edge_trace, node_trace],
                    layout=go.Layout(
                        title="Automation",
                        titlefont_size=8,
                        showlegend=False,
                        hovermode='closest',
                        margin=dict(b=3, l=5, r=0, t=0),
                        xaxis=dict(
                            showgrid=False,
                            zeroline=False,
                            showticklabels=False,
                            range=[-0.5, 10]  # x軸の表示範囲を制限
                        ),
                        yaxis=dict(
                            showgrid=False,
                            zeroline=False,
                            showticklabels=False,
                            range=[-1, 1]  # y軸の表示範囲を制限
                        ),
                        dragmode='pan'
                    )
                    )
    return fig

# アプリケーションのレイアウト
def create_layout(app):
    return html.Div([
    html.Div([
        html.Div(generate_info_content(df_info), style={'background-color': '#CDE1FF', 'font-size': '8px', 'padding': '5px', 'margin-bottom': '0px', 'flex': '3'}),
        html.Div([html.Img(src=image_url, style={'width': '90%', 'height': '60%', 'margin-bottom': '1px'})], style={'flex': '1', 'padding': '2px'})
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

    html.H1("Access Control & Automation", style={'textAlign': 'center'}),
    dcc.Graph(
        id='network-graph-access',
        figure=create_access_control_graph(),
        style={'height': '200px', 'width': '100%'}
    ),
    dcc.Graph(
        id='network-graph-automation',
        figure=create_automation_graph(),
        style={'height': '450px', 'width': '210%'}
    ),
    html.Div(id='click-output', style={'textAlign': 'center'}),
    
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

# コールバックでノードクリック時の処理を実装
@app.callback(
    Output('click-output', 'children'),
    [Input('network-graph-access', 'clickData'),
     Input('network-graph-automation', 'clickData')]
)
def display_click_data(clickData_access, clickData_automation):
    if clickData_access is not None:
        clicked_node = clickData_access['points'][0]['text']
        if clicked_node in nodes_access and nodes_access[clicked_node]["link"]:
            webbrowser.open(nodes_access[clicked_node]["link"])  # リンクをブラウザで開く
            return f"{clicked_node} のリンクを開きました: {nodes_access[clicked_node]['link']}"
    elif clickData_automation is not None:
        clicked_node = clickData_automation['points'][0]['text']
        if clicked_node in nodes_automation and nodes_automation[clicked_node]["link"]:
            webbrowser.open(nodes_automation[clicked_node]["link"])  # リンクをブラウザで開く
            return f"{clicked_node} のリンクを開きました: {nodes_automation[clicked_node]['link']}"

    return ""

# サーバーを実行
if __name__ == '__main__':
    app.run_server(debug=True)