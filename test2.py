import os
import xlrd
import xlwt
dirlist=os.listdir("./data")
print(dirlist)
class Guo():
    def __init__(self,dir_dict):
        self.dir_dict=dir_dict
    def guoming(self):
        try:
            return dir_dict['国名']
        except:
            return ""
    def mianji(self):
        try:
            return dir_dict['面积']
        except:
            return ""
    def renkou(self):
        try:
            return dir_dict['人口']
        except:
            return ""
    def guanfangyuyan(self):
        try:
            return dir_dict['官方语言']
        except:
            return ""
    def jiankuang(self):
        try:
            return dir_dict['简况']
        except:
            return ""
    def zhengzhi(self):
        try:
            return dir_dict['政治']
        except:
            return ""
    def xingzhengquhua(self):
        try:
            return dir_dict['行政区划']
        except:
            return ""
    def jingji(self):
        try:
            return dir_dict['经济']
        except:
            return ""
    def ziyuan(self):
        try:
            return dir_dict['资源']
        except:
            return ""
    def gongye(self):
        try:
            return dir_dict['工业']
        except:
            return ""
    def nongye(self):
        try:
            return dir_dict['农业']
        except:
            return ""
    def jiaotongyunshu(self):
        try:
            return dir_dict['交通运输']
        except:
            return ""
    def caizhengjinrong(self):
        try:
            return dir_dict['财政金融']
        except:
            return ""
    def duiwaimaoyi(self):
        try:
            return dir_dict['对外贸易']
        except:
            return ""
    def renmshenghuo(self):
        try:
            return dir_dict['人民生活']
        except:
            return ""
    def junshi(self):
        try:
            return dir_dict['军事']
        except:
            return ""
    def jiaoyu(self):
        try:
            return dir_dict['教育']
        except:
            return ""
    def xinwenchuban(self):
        try:
            return dir_dict['新闻出版']
        except:
            return ""
    def duiwaiguanxi(self):
        try:
            return dir_dict['对外关系']
        except:
            return ""

def download(guo,dir):
    worksheet.write(dirlist.index(dir), 0, label=guo.guoming())
    worksheet.write(dirlist.index(dir), 1, label=guo.mianji())
    worksheet.write(dirlist.index(dir), 2, label=guo.renkou())
    worksheet.write(dirlist.index(dir), 3, label=guo.guanfangyuyan())
    worksheet.write(dirlist.index(dir), 4, label=guo.jiankuang())
    worksheet.write(dirlist.index(dir), 5, label=guo.zhengzhi())
    worksheet.write(dirlist.index(dir), 6, label=guo.xingzhengquhua())
    worksheet.write(dirlist.index(dir), 7, label=guo.jingji())
    worksheet.write(dirlist.index(dir), 8, label=guo.ziyuan())
    worksheet.write(dirlist.index(dir), 9, label=guo.gongye())
    worksheet.write(dirlist.index(dir), 10, label=guo.nongye())
    worksheet.write(dirlist.index(dir), 11, label=guo.jiaotongyunshu())
    worksheet.write(dirlist.index(dir), 12, label=guo.caizhengjinrong())
    worksheet.write(dirlist.index(dir), 13, label=guo.duiwaimaoyi())
    worksheet.write(dirlist.index(dir), 14, label=guo.renmshenghuo())
    worksheet.write(dirlist.index(dir), 15, label=guo.junshi())
    worksheet.write(dirlist.index(dir), 16, label=guo.jiaoyu())
    worksheet.write(dirlist.index(dir), 17, label=guo.xinwenchuban())
    worksheet.write(dirlist.index(dir), 18, label=guo.duiwaiguanxi())
    worksheet.write(dirlist.index(dir), 19, label=dir)
    workbook.save("lalala.xls")


if __name__ == '__main__':
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('1')
    for dir in dirlist:
        data_dir=xlrd.open_workbook(r"./data/"+dir)
        table = data_dir.sheet_by_index(0)
        dir_key = table.col_values(0, 1)
        dir_key=[i.replace(" ", "") for i in dir_key]
        dir_value = table.col_values(1, 1)
        dir_dict=dict(zip(dir_key,dir_value))
        print(dir_dict)
        guo=Guo(dir_dict)
        download(guo,dir)


        # https___www.fmprc.gov.cn_web_gjhdq_676201_gj_676203_bmz_679954_1206_680156_1206x0_680158_.xls
        # https://www.fmprc.gov.cn/web/gjhdq_676201/gj_676203/bmz_679954/1206_680156/1206x0_680158/