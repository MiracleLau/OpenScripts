import requests
import json
from urllib.parse import unquote_plus
import argparse

'''
腾讯企业邮箱通讯录导出工具
未做任何容错，请正确设置各项值
'''

class QQExmailExport:
    _url = "https://exmail.qq.com/cgi-bin/laddr_lastlist?sid=&action=show_all_fast&t=addr_data_fast.json&f=js&ef=js&type=list_recent_contact&responseStatus=new&scode=&ignore_bizuser="
    _header = {
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36 Edg/87.0.664.60",
        "Origin":"https://exmail.qq.com",
        "Referer": "https://exmail.qq.com/zh_CN/htmledition/domain2.html?"
    }
    _mailJson = ""

    # 设置Cookie，获取通信录前需要先调用该方法
    def set_cookie(self,cookie):
        self._header["Cookie"] = cookie
    
    # 获取通信录的Json信息并解析
    def get_json(self):
        print("请求接口")
        re = requests.get(self._url,headers=self._header)
        jsonText = unquote_plus(re.text.replace('\\', '\\\\'))
        print("解析数据")
        self._mailJson = json.loads(jsonText, strict=False)

    # 获取邮箱群组
    def get_email_group(self):
        groupFile = open("groups.csv","a+")
        groupFile.write("项目组名称,邮箱地址\n")
        emailInfo = self._mailJson["emailgrouplist"]
        print("写入邮箱群组信息")
        for i in emailInfo:
            groupFile.write("%s,%s\n"%(i["name"].replace(u'\xa0', u' '),i["email"].replace(u'\xa0', u' ')))
        groupFile.close()
        print("写入成功")

    # 获取邮箱列表，包括个人通信录和企业通信录里的所有的联系人
    def get_mail(self):
        mailFile = open("mail.csv","a+")
        mailFile.write("姓名,邮箱,所属组织\n")
        mailInfo = self._mailJson["addrsets"]
        print("写入联系人信息")
        for i in mailInfo:
            for j in mailInfo[i]:
                mailFile.write("%s,%s,%s\n"%(j["name"].replace(u'\xa0', u' '),j["email"].replace(u'\xa0', u' '),j["department"].replace(u'\xa0', u' ')))
        mailFile.close()
        print("写入成功")

    # 运行
    def run(self):
        self.get_json()
        self.get_email_group()
        self.get_mail()

if __name__=="__main__":
    parser = argparse.ArgumentParser(description='腾讯企业邮箱通讯录导出工具')
    parser.add_argument('-s', required=True, help='Cookie中qm_sid的值')
    parser.add_argument('-u', required=True, help='Cookie中的biz_username或qm_username的值')
    args = parser.parse_args()
    cookie = "qm_sid=%s; biz_username=%s; qm_username=%s"%(args.s,args.u,args.u)
    m = QQExmailExport()
    m.set_cookie(cookie)
    m.run()
