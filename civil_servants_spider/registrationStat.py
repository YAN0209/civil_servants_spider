#!/usr/bin/python
# _*_ coding:utf-8 _*_
"""
  @author: likaiyan
  @date: 2021/2/5 上午11:22
  @desc:
"""

import argparse
import time
import requests
import json
import xlwt


class RegistrationStat:

    def __init__(self):
        cookies, target_path = self.parse_args()
        self.cookies = cookies
        self.target_path = target_path

    def parse_args(self):
        default_file_name = "报名统计详情" + time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime()) + ".xls"
        parser = argparse.ArgumentParser()
        parser.add_argument("cookies")
        parser.add_argument("--target_path", default=default_file_name)
        args = parser.parse_args()
        return args.cookies, args.target_path

    def get_headers(self):
        headers = {
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": self.cookies,
            "Host": "ggfw.gdhrss.gov.cn",
            "Origin": "https://ggfw.gdhrss.gov.cn",
            "Referer": "https://ggfw.gdhrss.gov.cn/gwyks/center.do?nvt=1612493336699",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0",
            "X-Requested-With": "XMLHttpRequest"
        }
        return headers

    def create_excel(self, datas):
        wb = xlwt.Workbook()

        sheet = wb.add_sheet("报名统计详情")

        sheet.write(0, 0, "招考单位")
        sheet.write(0, 1, "招考职位")
        sheet.write(0, 2, "职位代码")
        sheet.write(0, 3, "成功缴费人数")

        i = 1
        for row in datas:
            sheet.write(i, 0, row["aab004"])
            sheet.write(i, 1, row["bfe3a4"])
            sheet.write(i, 2, row["bfe301"])
            sheet.write(i, 3, row["aab119"])
            i = i + 1

        wb.save(self.target_path)

    def run(self):
        headers = self.get_headers()

        datas = []

        for i in range(1, 21):
            print("爬取第{}页".format(i))
            data = {
                "bfa001": "202101",
                "bab301": "01",
                "page": i,
                "rows": "50"
            }
            r = requests.post("https://ggfw.gdhrss.gov.cn/gwyks/exam/details/spQuery.do", data=data, headers=headers)
            if r.status_code == 200:
                response_data = json.loads(r.text)
                datas += response_data["rows"]
            else:
                print("获取第{}页数据失败".format(data["page"]))
            time.sleep(3)

        self.create_excel(datas)



if __name__ == '__main__':
    RegistrationStat().run()
