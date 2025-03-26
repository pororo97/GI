import dash
from dash import dcc, html, Input, Output
import pandas as pd
import plotly.express as px

# 初始化Dash应用
app = dash.Dash(__name__)

# Excel文件路径（示例路径，实际使用时请替换）
excel_file_path = r"C:\Users\1240509\Desktop\グローバルインテリジェンス\競合会社業績.xlsx"

# 公司和年份列表
companies = ["Assa", "dormakaba", "Allegion", "Sanwa", "NTS"]
years = [2019, 2020, 2021, 2022, 2023]

# 指标名称
metrics = [
    "Sales(売上)",
    "Cost of goods sold%(売上原価)",
    "SGA%(販管費)",
    "R&D%(研究開発費)",
    "EBIT％",
    "Entrance%",
    "Organic Growth",
]

# Dash布局
def create_layout(app):
    return html.Div([
    # 页面标题
    html.H1(
        "競合会社業績比較",
        style={
            'textAlign': 'center',
            'fontWeight': 'bold',
            'fontSize': '24px',
            'marginBottom': '20px'
        }
    ),
    # 主内容区域（包含筛选栏和图表）
    html.Div([
        # 左侧筛选栏和注释
        html.Div([
            html.Div("会社選択", style={
                'fontWeight': 'bold',
                'fontSize': '16px',
                'marginBottom': '10px'
            }),
            html.Div([
                dcc.Checklist(
                    id='company-selector',
                    options=[{'label': company, 'value': company} for company in companies],
                    value=companies,
                    style={
                        'display': 'block',
                        'marginBottom': '5px',
                        'backgroundColor': 'rgb(219, 228, 245)',
                        'border': '1px solid black',
                        'color': 'black',
                        'padding': '5px',
                        'borderRadius': '5px',
                    },
                    inputStyle={"marginRight": "10px"}
                )
            ], style={'marginBottom': '20px'}),
            html.Div("年", style={
                'fontWeight': 'bold',
                'fontSize': '16px',
                'marginBottom': '10px'
            }),
            html.Div([
                dcc.Checklist(
                    id='year-selector',
                    options=[{'label': year, 'value': year} for year in years],
                    value=years,
                    style={
                        'display': 'block',
                        'marginBottom': '5px',
                        'backgroundColor': 'rgb(219, 228, 245)',
                        'border': '1px solid black',
                        'color': 'black',
                        'padding': '5px',
                        'borderRadius': '5px',
                    },
                    inputStyle={"marginRight": "10px"}
                )
            ]),
            # 注释文本框
            html.Div([
                html.P("注: Sales(売上)は日本円換算", style={'marginBottom': '10px'}),
                html.P("    dormakabaは6月計算のため、2019年度は2018年7月ー2019年6月の業績;Sanwaは4月決算のため、2019年度は2019年4月ー2020年3月の業績", style={'marginBottom': '8px'}),
                html.P("Entrance％:", style={'marginBottom': '5px'}),
                html.P("  Assa: ENTRANCE SYSTEMの全社対売上シェア", style={'marginBottom': '5px'}),
                html.P("  dormakaba: Access Solutionsの全社対売上シェア", style={'marginBottom': '5px'}),
                html.P("  Allegion: 2021年以前はdoors/door syetemsの全社対売上シェア; 2021年以降はdoors/accessories/otherの全社対売上シェア", style={'marginBottom': '5px'}),
                html.P("  Sanwa: 米州歩行者の全社対売上シェア", style={'marginBottom': '5px'}),
                html.P("  NTS: AIGの全社対売上シェア", style={'marginBottom': '5px'}),
            ], style={
                'marginTop': '20px',
                'padding': '10px',
                'backgroundColor': 'rgb(240, 240, 240)',
                'border': '1px solid black',
                'borderRadius': '5px',
                'fontSize': '14px',
                'lineHeight': '1.5'
            })
        ], style={
            'width': '20%',
            'padding': '10px',
            'backgroundColor': 'rgb(219, 228, 245)',
            'border': '1px solid black',
            'borderRadius': '10px',
            'boxShadow': '2px 2px 5px rgba(0, 0, 0, 0.1)'
        }),
    
        # 右侧图表区域
        html.Div([
            html.Div([
                dcc.Graph(id=f'graph-{i}', style={'height': '300px'})  # 高度调整为300px
                for i in range(len(metrics))
            ], style={
                'display': 'grid',
                'gridTemplateColumns': '1fr 1fr 1fr',  # 3列布局
                'gap': '5px',  # 图表之间的间隙调整为10px
                'padding': '2px'  # 外部间隙调整为5px
            }),
        ], style={'width': '85%', 'padding': '10px'})
    ], style={
        'display': 'flex',  # 使用flex布局实现水平对齐
        'alignItems': 'flex-start',  # 垂直方向顶部对齐
        'justifyContent': 'space-between'  # 两部分之间留一定间距
    }),
    html.A("目次戻る", href="/", style={
            "position": "absolute", "top": "10px", "right": "10px",
            "backgroundColor": "#f0f0f0", "padding": "5px",
            "borderRadius": "5px", "textDecoration": "none"
        })
])

# 回调函数：根据选择动态更新图表
@app.callback(
    [Output(f'graph-{i}', 'figure') for i in range(len(metrics))],
    [Input('company-selector', 'value'), Input('year-selector', 'value')]
)
def update_graphs(selected_companies, selected_years):
    figures = []
    for i, metric in enumerate(metrics):
        # 初始化空数据框
        combined_data = pd.DataFrame()
        # 遍历每个年份
        for year in selected_years:
            sheet_name = str(year)
            data = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="A:F")
            data = data.rename(columns={data.columns[0]: "Metric"})
            data = data[data["Metric"] == metric]
            data = data.melt(id_vars="Metric", var_name="Company", value_name="Value")
            data["Year"] = year
            combined_data = pd.concat([combined_data, data])
        # 筛选选中的公司
        combined_data = combined_data[combined_data["Company"].isin(selected_companies)]
        # 创建图表
        fig = px.line(
            combined_data,
            x="Year",
            y="Value",
            color="Company",
            title=metric,
            markers=True
        )
        # 更新图表布局
        fig.update_layout(
            yaxis_title="",
            xaxis_title="",
            legend_title_text="",  # 设置凡例标题为空
            legend_orientation="h",  # 横向显示
            legend_y=-0.3,           # 凡例在图表下方
            legend_x=0.5,            # 凡例居中
            legend_xanchor="center",  # 凡例的中心对齐
            margin=dict(l=10, r=10, t=30, b=10)  # 图表的内边距
        )
        # 根据指标设置Y轴格式
        if metric != "Sales(売上)":
            fig.update_layout(yaxis_tickformat=".1%")  # 百分比格式，保留一位小数
        figures.append(fig)
    return figures

# 运行应用
if __name__ == '__main__':
    app.run_server(debug=True)