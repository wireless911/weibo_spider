# coding=utf-8
import urllib3
import requests
from bs4 import BeautifulSoup
import re
from xlwt import *
import xlrd
import os
from xlutils.copy import copy
import time
import json

index = 100


class Spider(object):
    """实现微博品牌kol图片、微博内容爬虫

    If the class has public attributes, they may be documented here
    in an ``Attributes`` section and follow the same formatting as a
    function's ``Args`` section. Alternatively, attributes may be documented
    inline with the attribute's declaration (see __init__ method below).

    Properties created with the ``@property`` decorator should be documented
    in the property's getter method.

    Attributes:
		cookie: 	A string of user's cookie　
		pages: 	    A integer of dog's name
		profile: 	A string of brand's name
		sleep_time:  	A integer of intervals


	pg:
	    cookie 有过期时间，请求失败时更换cookie

    """

    def __init__(self, cookie, pages=1, profile="perfectdiary", sleep_time=15):
        self.sleep_time = sleep_time
        self.pages = pages
        self.cookie = cookie
        self.profile = profile
        self.script_uri = "/" + profile
        self._iter_page()

    def _iter_page(self):
        for page in range(1, self.pages):
            print(time.strftime('%Y-%m-%d %X', time.localtime()))
            result0 = self.get_response(page)
            self.save_data(result0)
            time.sleep(self.sleep_time)

    def get_response(self, page):
        """微博每页的数据分三次请求，初始页面为js渲染html，下拉请求json数据渲染，需要拼接参数"""
        requests.packages.urllib3.disable_warnings()
        http = urllib3.PoolManager()
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.108 Safari/537.36',
            'Cookie': self.cookie,
            "X-Requested-With": "XMLHttpRequest"
        }
        start_url = 'https://weibo.com/{profile}?pids=Pl_Official_MyProfileFeed__23&is_search=0&visible=0&is_hot=1&is_tag=0&profile_ftype=1&page={page}&ajaxpagelet=1&ajaxpagelet_v6=1&__ref=%2Fperfectdiary%3Fis_search%3D0%26visible%3D0%26is_hot%3D1%26is_tag%3D0%26profile_ftype%3D1%26page%3D3%23feedtop&_t=FM_157441560856733'.format(
            profile=self.profile, page=page)

        r = http.request('GET', start_url, headers=headers)
        data = json.loads(r.data.decode().strip()[23:-10]).get("html")
        soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

        result0 = soup.find_all("div", attrs={"action-type": "feed_list_item"})

        for pagebar in [0, 1]:
            json_url = "https://weibo.com/p/aj/v6/mblog/mbloglist?ajwvr=6&domain=100606&is_search=0&visible=0&is_all=1&is_tag=0&profile_ftype=1&page={page}&pagebar={pagebar}&pl_name=Pl_Official_MyProfileFeed__23&id=1006066020329578&script_uri={script_uri}&feed_type=0&pre_page={pre_page}&domain_op=100606&__rnd=1575859271326".format(
                page=page, pagebar=pagebar, pre_page=page, script_uri=self.script_uri)
            res = http.request('GET', json_url, headers=headers)
            json_data = json.loads(res.data.decode().strip()).get("data")
            json_soup = BeautifulSoup(json_data, 'html.parser', from_encoding='utf-8')
            result0 += json_soup.find_all("div", attrs={"action-type": "feed_list_item"})
        return result0

    def save_data(self, result0):
        """图片保存在image文件夹中，微博详情保存在excel文件中，图片id对应微博在excel中"""
        global index
        i = 0
        dir_path = r'image/'
        excel_list = []
        files = os.listdir("./")
        for file in files:
            if file.endswith('weibo.xls') and '~$' not in file:
                excel_list.append(os.path.join("./", file))
        for excel in excel_list:
            workbook = xlrd.open_workbook(excel)
            sheet = workbook.sheet_by_index(0)
            i = sheet.nrows
            all_rows = [sheet.row_values(i) for i in range(sheet.nrows)]
            x = {row[0]: row[1] for row in all_rows}
            print(i)
        if excel_list:
            workbookX = Workbook(encoding='utf-8')
            workbookX = xlrd.open_workbook(os.path.join("./", 'weibo.xls'))
            workbookZ = copy(workbookX)
            sheetZ = workbookZ.get_sheet(0)

            for line0 in result0:
                mmid = 0
                mid = line0.get("mid")

                imgList = line0.find("div", class_="WB_detail").find_all("img")
                imgList = [i["src"] for i in imgList]
                imgList = [i for i in imgList if i.endswith(".jpg")]

                imgName = 0
                for imgPath in imgList:
                    try:
                        if not imgPath.startswith('https'):
                            imgPath = 'http:' + imgPath
                        print(imgPath)
                        if "orj360" in imgPath:
                            imgPath = imgPath.replace("orj360", "mw690")
                        elif "thumb150" in imgPath:
                            imgPath = imgPath.replace("thumb150", "mw690")
                        print(imgPath, "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                        http = urllib3.PoolManager()
                        img0 = http.request('GET', imgPath)
                        imgPath = os.path.join(dir_path, str(index) + "-" + str(imgName) + '.jpg')
                        with open(imgPath, 'wb') as imgF:
                            imgF.write(img0.data)
                            mmid += 1
                    except Exception as e:
                        print(e, "-------------------------------------------------------------------------")
                        print(imgPath + " error")
                    imgName = imgName + 1

                resultIn = line0.find("div", attrs={"node-type": "feed_list_content"}).get_text().strip()
                info = line0.find("a", class_="S_txt2").get_text() if line0.find("a", class_="S_txt2") else None
                sheetZ.write(i, 0, index)
                sheetZ.write(i, 1, str(resultIn))
                sheetZ.write(i, 2, time.strftime('%Y-%m-%d %X', time.localtime()))
                sheetZ.write(i, 3, info)
                print(mid)
                i = i + 1
                index += 1
        workbookZ.save(os.path.join("./", 'weibo.xls'))


if __name__ == '__main__':
    cookie = "SINAGLOBAL=6708852595869.952.1569471246335; wb_timefeed_3805868427=1; Ugrow-G0=140ad66ad7317901fc818d7fd7743564; login_sid_t=bd902682bf576e352a773eec63c7c4e1; cross_origin_proto=SSL; YF-V5-G0=f5a079faba115a1547149ae0d48383dc; WBStorage=42212210b087ca50|undefined; _s_tentry=passport.weibo.com; Apache=6855530219190.078.1575628001385; ULV=1575628001402:9:1:1:6855530219190.078.1575628001385:1574910237723; wb_view_log=1536*8641.25; crossidccode=CODE-gz-1IDaOU-28JYu8-1pylnaAkkBxWAaKe8a6f3; ALF=1607164028; SSOLoginState=1575628029; SCF=An2tHEqaeho57ne5Ti5N7lU0Xf7iuVwizJ-n25DxonO-yoC7ih-BWZeSYBQaF7pz5Z0KNUa3SKXEfcvKMxiYYQ0.; SUB=_2A25w7lyuDeRhGeVG61cZ9ibIyTuIHXVTmslmrDV8PUNbmtBeLUb4kW9NT9LFvUuoVg_GhVuqsgV64xbDrSQEXCcQ; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WFCX6lXRHpCC0ucIHyVs3Yu5JpX5KzhUgL.FoeReh-RSonXeoM2dJLoI0YLxKqL1-BLBK-LxKnL1hMLB-2LxKBLBonLBo.LxK.LB.2LBKqLxKML1-2L1hBLxK-LBKBLBK.LxKMLBo2L1h2t; SUHB=0U1J-AfpAEm67c; wvr=6; UOR=www.cnblogs.com,open.weibo.com,graph.qq.com; wb_view_log_3805868427=1536*8641.25; YF-Page-G0=02467fca7cf40a590c28b8459d93fb95|1575628059|1575628035; webim_unReadCount=%7B%22time%22%3A1575628225555%2C%22dm_pub_total%22%3A0%2C%22chat_group_client%22%3A0%2C%22allcountNum%22%3A0%2C%22msgbox%22%3A0%7D"

    sp = Spider(cookie)
