import requests
from bs4 import BeautifulSoup
import re
import pandas as pd
from pandas.core.frame import DataFrame

from Bilibili_Rank import analyse


def getHTMLText(url):
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        return ''


def get_title(html):
    title = re.search(r'统计所有投稿在 .*? 的数据综合得分', html).group(0)
    title_str = str(title)
    sheet_title = title_str[8:-8]
    return sheet_title


def Html_Analyse(html, AV_info_List):
    soup = BeautifulSoup(html, 'lxml')

    Av_rank = soup.select('.num')
    rank_list = []
    for i in Av_rank:
        rank = i.string
        rank_list.append(rank)

    Info = soup.select('script')
    Info_str_x = str(Info[5])
    Info_str = re.sub(r'"others":\[.*?\]', '', Info_str_x)  # 去除B站推荐内容
    author_list = re.findall(r'"author":".*?"', Info_str)  # UP主名字初始列表
    coins_list = re.findall(r'"coins":\d*', Info_str)  # 投币数量初始列表
    playnums_list = re.findall(r'"play":\d*', Info_str)  # 播放数初始列表
    pts_list = re.findall(r'"pts":\d*', Info_str)  # 综合评分初始列表
    title_list = re.findall(r'"title":".*?"', Info_str)  # 视频名字初始列表
    video_review_list = re.findall(r'"video_review":\d*', Info_str)  # 视频弹幕总数初始列表
    aid_list = re.findall(r'"aid":"\d*"', Info_str)  # 视频AV号初始列表

    '''
    以下用于生成每个视频的信息列表
    列表中每个元素中包含的信息顺序为：
    [排名，AV号，UP名，标题，综合评分，总播放量，投币数量，弹幕总数]
    '''

    for i in range(len(rank_list)):
        AV_info_List.append([])
        AV_info_List[i].append(rank_list[i])
        AV_info_List[i].append('AV'+aid_list[i][7:-1])
        AV_info_List[i].append(author_list[i][10:-1])
        AV_info_List[i].append(title_list[i][9:-1])
        AV_info_List[i].append(pts_list[i][6:])
        AV_info_List[i].append(playnums_list[i][7:])
        AV_info_List[i].append(coins_list[i][8:])
        AV_info_List[i].append(video_review_list[i][15:])

    return AV_info_List


def excel_save(List, sheet_title):
    df = DataFrame(List)
    df.columns = ['排名', 'AV号', 'UP名', '标题', '综合评分', '总播放量', '投币数量', '弹幕总数']
    writer = pd.ExcelWriter('B站{0}综合排行榜前100视频.xlsx'.format(sheet_title))
    df.to_excel(excel_writer=writer, index=False, encoding='utf-8', sheet_name=sheet_title)
    writer.save()
    writer.close()


def main():
    url = 'https://www.bilibili.com/ranking/all/0/0/1'
    AV_info_List = []  # 排行榜信息列表
    html = getHTMLText(url)
    sheet_title = get_title(html)
    List = Html_Analyse(html, AV_info_List)
    excel_save(List, sheet_title)
    analyse(sheet_title)


main()
