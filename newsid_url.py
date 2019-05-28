import xlrd
import xlwt
import json
import urllib.request
import random
import time

class newsid_to_url():

    def read_style(self):
    	#读取excel
        excel = xlrd.open_workbook('D:/日报/newsid.xlsx')
        table = excel.sheet_by_index(0)
    

        style = xlwt.XFStyle()
        #设置字体大小
        font = xlwt.Font()
        font.name = '微软雅黑'
        font.bold = False
        font.height = 200
        style.font = font
        
        #设置单元格框线
        borders = xlwt.Borders()
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        style.borders = borders

        #设置单元格内容位置
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_LEFT
        alignment.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = alignment
        
        #设置表头字体
        style1 = xlwt.XFStyle()
        alignment1 = xlwt.Alignment()
        alignment1.horz = xlwt.Alignment.HORZ_CENTER
        alignment1.vert = xlwt.Alignment.VERT_CENTER
        style1.alignment = alignment1
   
    	# 创建一个新的workbook 设置编码
        workbook = xlwt.Workbook(encoding = 'utf-8')

        # 创建一个worksheet
        worksheet = workbook.add_sheet('url')
        worksheet.write(0,0,'newsid',style1)
        worksheet.write(0,1,'url',style1)        
        
        for n in range(1,table.nrows):
            #读取excel中的newisd
            newsid = table.cell(n, 0).value
            data = {'newsid': newsid}

            #打开每个newsid的接口
            html = urllib.request.urlopen(r'****?newsId={}'.format(data['newsid']))
            #取出接口中的json
            hjson = json.loads(bytes.decode(html.read()))
           
            #将url与newsid写入excel中
            worksheet.write(n,1,hjson['url'],style)
            worksheet.write(n,0,data['newsid'],style)
        
        # 保存，名字取随机值
        name = time.strftime("%Y%m%d%H%M%S")
        print(name)
        workbook.save('D:/日报/url-{}.xls'.format(name))


if __name__ == '__main__':
    a = newsid_to_url().read_style()