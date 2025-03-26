from flask import Flask
from dash import dcc, html, Input, Output, Dash
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from apps.Assa.app import create_assa_app
from apps.dormakaba.app import create_dormakaba_app
from apps.FAAC.app import create_faac_app
from apps.Geze.app import create_geze_app
from apps.GU.app import create_gu_app
from apps.Horton.app import create_horton_app
from apps.Manusa.app import create_manusa_app
from apps.NewsSearch.app import create_newssearch_app

# Flaskサーバーのインスタンスを作成
server1 = Flask(__name__)


# 各アプリのインスタンスを作成
apps = {
    'Assa': create_assa_app(server1),
    'dormakaba': create_dormakaba_app(server1),
    'FAAC': create_faac_app(server1),
    'Geze': create_geze_app(server1),
    'GU': create_gu_app(server1),
    'Horton': create_horton_app(server1),
    'Manusa': create_manusa_app(server1),
    'NewsSearch': create_newssearch_app(server1),
}

# メインアプリのレイアウト
app = Dash(__name__, server=server1, name='main_app')
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Nav([
        html.Ul([
            html.Li(html.A(name, href=f'/{name}/')) for name in apps.keys()
        ])
    ]),
    html.Div(id='page-content')
])

# ページ遷移のコールバック
@app.callback(
    Output('page-content', 'children'),
    Input('url', 'pathname'))
def display_page(pathname):
    if pathname and pathname.strip('/') in apps:
        return apps[pathname.strip('/')].layout
    return '404 Page Not Found'

# WSGIアプリケーションとしての設定
application = DispatcherMiddleware(server1, {f'/{name}/': app for name in apps.keys()})

if __name__ == '__main__':
    from werkzeug.serving import run_simple
    run_simple('127.0.0.1', 8050, application, use_reloader=True, use_debugger=True)