import pandas as pd
from dash import Dash, dcc, html, Input, Output

# Excelデータの読み込み
file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\news.xlsx"
df = pd.read_excel(file_path)

# 日時の形式を変更し、エラーハンドリングを追加
df['日時'] = pd.to_datetime(df['日時'], errors='coerce').dt.strftime('%Y-%m-%d')

# NaTを含む行を削除する場合（必要に応じて）
df = df.dropna(subset=['日時'])

# Dashアプリケーションの作成
app = Dash(__name__)

# アプリケーションのレイアウト
def create_layout(app):
    return html.Div(style={'display': 'flex', 'justify-content': 'space-between', 'padding': '20px'}, children=[
    html.Div(style={'flex': '1', 'margin-right': '20px'}, children=[
        html.H1("ニュース一覧"),
        # 会社名を選択するためのドロップダウン
        dcc.Dropdown(
            id='company-dropdown',
            options=[{'label': company, 'value': company} for company in df['会社'].unique()],
            placeholder='会社を選択してください'
        ),
        # キーワード検索用のテキストボックス
        dcc.Input(
            id='keyword-input',
            type='text',
            placeholder='キーワードを入力してください',
            style={'backgroundColor': 'rgb(121, 175, 147)', 'width': '95%', 'padding': '10px'}
        ),
    ]),
    html.Div(style={'flex': '2'}, id='news-output'),
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
])

# コールバック関数の定義
@app.callback(
    Output('news-output', 'children'),
    [Input('company-dropdown', 'value'),
     Input('keyword-input', 'value')]
)
def update_output(selected_company, keyword):
    # 初期状態で全データを表示
    filtered_df = df

    # 会社が選択された場合、データをフィルタリング
    if selected_company:
        filtered_df = filtered_df[filtered_df['会社'] == selected_company]

    # キーワードが入力されている場合は、さらにフィルタリング
    if keyword:
        filtered_df = filtered_df[filtered_df['タイトル'].str.contains(keyword, na=False) |
                                   filtered_df['和訳'].str.contains(keyword, na=False) |
                                   filtered_df['まとめ'].str.contains(keyword, na=False)]

    # 日付でソート（最新から古い順）
    filtered_df = filtered_df.sort_values(by='日時', ascending=False)

    news_list = []

    # フィルタリングされたニュースをリストに追加
    for _, row in filtered_df.iterrows():
        news_list.append(html.Div([
            html.H3(row['タイトル']),
            html.P(f"日時: {row['日時']}"),
            html.P(f"和訳: {row['和訳']}"),
            html.P(f"まとめ: {row['まとめ']}"),
            html.A("続きを読む", href=row['URL'], target="_blank")
        ]))

    # フィルタリングされたニュースがない場合のメッセージ
    if not news_list:
        return "該当するニュースはありません。"

    return news_list

if __name__ == '__main__':
    app.run_server(debug=True)