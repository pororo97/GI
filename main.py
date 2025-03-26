from dash import Dash, html, dcc
from dash.dependencies import Input, Output
import Assa
import dormakaba
import FAAC
import Geze
import GU
import Horton
import Manusa
import Panasonic
import portalp
import Stanley
import Tormax
import News
import 競合会社比較

# メインアプリの作成
app = Dash(__name__, suppress_callback_exceptions=True)

# 目次ページのレイアウト
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content'),
])

# 目次ページの内容
index_page = html.Div([
    html.H1("目次"),
    html.Ul([
        html.Li(dcc.Link('Assa App', href='/assa')),
        html.Li(dcc.Link('Dormakaba App', href='/dormakaba')),
        html.Li(dcc.Link('FAAC App', href='/faac')),
        html.Li(dcc.Link('Geze App', href='/geze')),
        html.Li(dcc.Link('GU App', href='/gu')),
        html.Li(dcc.Link('Horton App', href='/horton')),
        html.Li(dcc.Link('Manusa App', href='/manusa')),
        html.Li(dcc.Link('News Search App', href='/news_search')),
        html.Li(dcc.Link('Panasonic App', href='/panasonic')),
        html.Li(dcc.Link('Portalp App', href='/portalp')),
        html.Li(dcc.Link('Stanley App', href='/stanley')),
        html.Li(dcc.Link('Tormax App', href='/tormax')),
        html.Li(dcc.Link('競合会社比較 App', href='/competitor_comparison')),
    ])
])

# URLに基づくルーティング
@app.callback(Output('page-content', 'children'),
              Input('url', 'pathname'))
def display_page(pathname):
    if pathname == '/assa':
        return Assa.layout
    elif pathname == '/dormakaba':
        return dormakaba.layout
    elif pathname == '/faac':
        return FAAC.layout
    elif pathname == '/geze':
        return Geze.layout
    elif pathname == '/gu':
        return GU.layout
    elif pathname == '/horton':
        return Horton.layout
    elif pathname == '/manusa':
        return Manusa.layout   
    elif pathname == '/news_search':
        return News.layout
    elif pathname == '/panasonic':
        return Panasonic.layout
    elif pathname == '/portalp':
        return portalp.layout
    elif pathname == '/stanley':
        return Stanley.layout
    elif pathname == '/tormax':
        return Tormax.layout
    elif pathname == '/competitor_comparison':
        return 競合会社比較.layout
    else:
        return index_page

# アプリの実行
if __name__ == '__main__':
    app.run_server(debug=True)