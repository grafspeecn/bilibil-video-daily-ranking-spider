import pandas as pd
from pyecharts.charts import Line, Bar
import pyecharts.options as opts


def analyse(sheet_title):
    data = pd.read_excel('B站{0}综合排行榜前100视频.xlsx'.format(sheet_title), sheet_name=sheet_title)
    x_data = data['AV号']
    y_data_pts = data['综合评分']
    y_data_playnums = list(data['总播放量']/10)
    y_data_coins = list(data['投币数量'])
    y_data_views = list(data['弹幕总数'])
    bar = (
        Bar(init_opts=opts.InitOpts(width="1600px", height="800px"))
        .add_xaxis(xaxis_data=x_data,)
        .add_yaxis('总播放量', yaxis_data=y_data_playnums)
        .add_yaxis('投币数量', yaxis_data=y_data_coins, stack=1)
        .add_yaxis('弹幕总数', yaxis_data=y_data_views, stack=1)
        .extend_axis(yaxis=opts.AxisOpts(
            name='综合评分', type_='value'
        ))
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(title_opts=opts.TitleOpts(title="B站{0}综合排行榜前100视频".format(sheet_title),
                                                   pos_left="center",
                                                   pos_top="bottom"),
                         tooltip_opts=opts.TooltipOpts(
                             is_show=True, trigger="axis", axis_pointer_type="cross"
                         ),
                         xaxis_opts=opts.AxisOpts(
                             type_="category",
                             axispointer_opts=opts.AxisPointerOpts(is_show=True, type_="shadow"),
                             axislabel_opts=opts.LabelOpts(rotate=-30)
                         ),
                         yaxis_opts=opts.AxisOpts(
                             name='总播放量*0.1/投币数量/弹幕总数',
                             type_='value'
                         ),
                         datazoom_opts=opts.DataZoomOpts()
                         )
    )
    line = (
        Line()
        .add_xaxis(xaxis_data=data['AV号'])
        .add_yaxis(series_name='综合评分', yaxis_index=1, y_axis=y_data_pts,
                   label_opts=opts.LabelOpts(is_show=False))
    )
    bar.overlap(line).render("B站{0}综合排行榜前100视频数据.html".format(sheet_title))


