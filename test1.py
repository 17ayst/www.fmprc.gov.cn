import xlwt
import re
from lxml.html import etree
import requests
import xlwt


# 获取url的HTML的txt
def gettext(url):
    global headers
    req = requests.get(url, headers=headers,verify=False)
    # html0 = etree.HTML(req.text)
    # print(req.text)
    return req.text

# 解析各州网址
def prase1(url0):
    html0=etree.HTML(gettext(url0))
    url1_list= html0.xpath('/html/body/div[1]/div[5]/div[2]/div/table/tr/td/a/@href')
    return url1_list

def prase2(txt):
    html0 = etree.HTML(txt)
    list1 = html0.xpath('//*[@id="content"]/p/text()')
    return list1

def download(guo_url):
    txt_lian=""
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('1')
    # txt = gettext("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/yz_676205/1206_676356/1206x0_676358/")
    txt = gettext(guo_url)
    # print(txt)
    list1 = prase2(txt)
    for i in list1:
        # print(i)
        txt_lian += i
    list2 = txt_lian.split("【")
    for index in range(1, len(list2)):
        # print("-----------"+a)
        print(list2[index])
        list3 = list2[index].split("】")
        print(list3)
        worksheet.write(index, 0, label=list3[0])
        worksheet.write(index, 1, label=list3[1])
    workbook.save(guo_url.replace("/", "_").replace(":","_")+".xls")


if __name__ == '__main__':
    url={}
    # url_0="https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/bmz_679954"
    headers={"Cookie": "_trs_uv=kl64rnr5_469_cgx9; _trs_ua_s_1=kl64rnr5_469_f2ki; _Jo0OQK=16A6D7DABECCCF3D1102D53D3F1AF27A47F6179A84AC94A4433BBB3C291BA91C3C646EDEF552F829F94C6607DE0EE3B4AE9BA4B22874F3FEDE68AB11D564537A6AF1ABD0D27BB1D74A358EDFE261B6ACD9C58EDFE261B6ACD9CA79E0CB27D2A5E77GJ1Z1ew==","User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36"}
    # url_list1_add=prase1(url_0)
    # url_list1=[url_0+i.replace(".","") for i in url_list1_add]
    url["./yz_676205/"] = prase1("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/yz_676205/")
    url["./fz_677316/"] = prase1("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/fz_677316/")
    url["./oz_678770/"] = prase1("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/oz_678770/")
    url["./bmz_679954/"] = prase1("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/bmz_679954/")
    url["./nmz_680924/"] = prase1("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/nmz_680924/")
    url["./dyz_681240/"] = prase1("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/dyz_681240/")
    print(url)
    for zhou in url:
        zhou_guo_list=url[zhou]
        for guo in zhou_guo_list:
            # guo_url="https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/"+zhou.replace("./","")+guo.replace("./","")+"1206x0_676209/"
            window_location_txt=gettext("https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/"+zhou.replace("./","")+guo.replace("./",""))
            window_location = re.findall('(window.location.href=")(.*)(")', window_location_txt)[0][1]
            guo_url = "https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/" + zhou.replace("./", "") + guo.replace("./", "")+window_location.replace("./", "")
            # print(gettext(guo_url))
            print(guo_url)
            download(guo_url)


