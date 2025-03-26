import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import importlib
from google.colab import output


# 创建主应用
app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server

# 定义子应用的路由
sub_apps = {
    "Assa": "Assa",
    "dormakaba": "dormakaba",
    "FAAC": "FAAC",
    "Geze": "Geze",
    "GU": "GU",
    "Horton": "Horton",
    "Manusa": "Manusa",
    "Panasonic": "Panasonic",
    "portalp": "portalp",
    "Stanley": "Stanley",
    "Tormax": "Tormax",
    "競合会社比較": "競合会社比較",
    "News Search": "News Search"
}

# 子应用模块的预加载
loaded_apps = {}
for app_id in sub_apps.keys():
    try:
        module = importlib.import_module(f"apps.{app_id}")
        if hasattr(module, "create_layout"):
            loaded_apps[app_id] = module.create_layout
        else:
            print(f"モジュール {app_id} に create_layout 関数がありません")
    except ModuleNotFoundError:
        print(f"モジュール {app_id} が見つかりません")

# 目次ページのレイアウト
def get_directory_layout():
    return html.Div([
        html.H1("グローバルインテリジェンス", style={"textAlign": "center"}),
        html.Ul([
            html.Li(html.A(app_name, href=f"/{app_id}"))
            for app_id, app_name in sub_apps.items()
        ])
    ])

# アプリケーションのレイアウト
app.layout = html.Div([
    dcc.Location(id="url", refresh=False),
    html.Div(id="page-content")
])

# URL に基づいてページを表示
@app.callback(
    Output("page-content", "children"),
    [Input("url", "pathname")]
)
def display_page(pathname):
    if pathname == "/":
        return get_directory_layout()
    elif pathname and pathname.strip("/") in loaded_apps:
        return loaded_apps[pathname.strip("/")](app)  # app を渡す
    else:
        return html.Div([
            html.H1("404 ページが見つかりません", style={"textAlign": "center"}),
            html.A("目次に戻る", href="/", style={"textAlign": "center"})
        ])

# アプリケーションを実行
if __name__ == "__main__":
    app.run_server(debug=False)