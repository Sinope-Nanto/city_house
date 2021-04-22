import openpyxl
from openpyxl.styles import Side, Border,PatternFill, colors, Font, Alignment
from openpyxl.drawing.image import Image
import time

class GenExcelReport:
    def __init__(self,url):
        self.edge = Side(style='thin', color='000000')
        self.border = Border(top=self.edge, bottom=self.edge, left=self.edge, right=self.edge)
        self.url = url
        self.report = openpyxl.Workbook()
    def create_report(self):
        indexsummary = self.report.active
        indexsummary.title = '指数汇总'
        totalradioplot = self.report.create_sheet()
        totalradioplot.title = '全国指数同比环比图'
        plotsimple = self.report.create_sheet()
        plotsimple.title = '作图（简版）'
        plotcomplex = self.report.create_sheet()
        plotcomplex.title = '作图（详细）'
        citysummary = self.report.create_sheet()
        citysummary.title = '城市汇总'
        citysort = self.report.create_sheet()
        citysort.title = '城市排序'
        cityindex = self.report.create_sheet()
        cityindex.title = '各城市指数'
        cityvolumn = self.report.create_sheet()
        cityvolumn.title = '各城市样本量'
        cityplot = self.report.create_sheet()
        cityplot.title = '各城市作图'
        return 
    def EndReport(self):
        self.report.save(self.url)

    def to_date(self, temp):
        year = int(temp)//12 + 6
        month = int(temp)%12 + 1
        return str(year).zfill(2) + '-' + str(month).zfill(2)
    
    def IndexSummary(self,**kwargs):
        # 参数列表:['volumn','east_volumn','mid_volumn','west_volumn',
        # 'under_90_volumn','90_144_volumn','above_144_volumn','volumn_year_on_year','volumn_chain',
        # 'index','east_index','mid_index','west_index',
        # 'year_on_year','chain','east_year_on_year','east_chain',
        # 'mid_year_on_year','mid_chain','west_year_on_year','west_chain',
        # 'index_under_90','index_90_144','index_above_144','year_on_year_under_90',
        # 'chain_under_90','year_on_year_90_144','chain_90_144','year_on_year_above_144','chain_above_144','index_CSJ']
        
        # 表格基本格式设置
        indexsummary = self.report['指数汇总']
        indexsummary.freeze_panes = 'C1'
        indexsummary.column_dimensions['A'].width = 20.0
        titlefont = Font(u'宋体',size = 12,bold=True,color='000000')
        datafont = Font(u'宋体',size = 10,bold=True,color='000000')
        grayfill = PatternFill("solid", fgColor='00C0C0C0')
        yellowfill = PatternFill("solid", fgColor='FFFFFF00')
        col_max = len(kwargs['volumn'])+2
        
        # 表头文字设置
        titlelist = ['交易量','指数汇总结果','区域性指数']
        rowlist = [1,15,43]
        for i in range(0,3):
            c = indexsummary.cell(row=rowlist[i],column=1)
            c.value = titlelist[i]
            c.font = titlefont
        
        # 表格颜色填充
        grayrow = [1,15,43]
        yellowrow = [24,25,28,29,33,38,39]
        for j in range(1,col_max):
            for i in grayrow:
                c = indexsummary.cell(row=i,column=j)
                c.fill = grayfill
            for i in yellowrow:
                c = indexsummary.cell(row=i,column=j)
                c.fill = yellowfill
        
        # 表格边框加粗
        borderrow = [i for i in range(2,10)] + [11,12] + [i for i in range(16,21)] + [i for i in range(22,30)]+ [31,32,33,34]+ [i for i in range(36,42)]+ [44,45]
        for i in borderrow:
            for j in range(1,col_max):
                c = indexsummary.cell(row=i,column=j)
                c.border = self.border

        # 日期填入
        datarow = [2,16,31,44]
        for j in datarow:
            for i in range(2,col_max):
                c = indexsummary.cell(row=j,column=i)
                c.value = self.to_date(i-2)
                c.font = datafont

        # 表头单元格合并
        indexsummary.merge_cells('A1:B1')
        indexsummary.merge_cells('A15:B15')
        indexsummary.merge_cells('A43:B43')

        # 填入表行名
        int_row_name = ['全国','东部','中部','西部','90以下','90-144','144以上']
        float_row_name = ['全国','东部','中部','西部','90以下','90-144','144以上','长三角']
        year_row_name = ['全国交易量同比变化','全国同比涨幅','东部同比涨幅','中部同比涨幅','西部同比涨幅',
        '90以下同比涨幅','90-144同比涨幅','144以上同比涨幅']
        chain_row_name = ['全国交易量环比变化','全国环比涨幅','东部环比涨幅','中部环比涨幅','西部环比涨幅',
        '90以下环比涨幅','90-144环比涨幅','144以上环比涨幅']
        int_row_list = [i for i in range(3,10)]
        float_row_list = [17,18,19,20,32,33,34,45]
        year_row_list = [11,22,24,26,28,36,38,40]
        chain_row_list = [12,23,25,27,29,37,39,41]

        rowname = int_row_name + float_row_name + year_row_name + chain_row_name
        rowlist = int_row_list + float_row_list + year_row_list + chain_row_list

        for i in range(0,len(rowname)):
            c = indexsummary.cell(row=rowlist[i],column=1)
            c.value = rowname[i]
            c.font = datafont


        # 填入整数型数据
        data_name = ['volumn','east_volumn','mid_volumn','west_volumn',
        'under_90_volumn','90_144_volumn','above_144_volumn']
        for i in range(0,len(int_row_list)):
            for j in range(2,col_max):
                c = indexsummary.cell(row=int_row_list[i],column=j)
                c.value = kwargs[data_name[i]][j-2]
                c.font = datafont

        # 填入浮点数型数据
        data_name = ['index','east_index','mid_index','west_index',
        'index_under_90','index_90_144','index_above_144','index_CSJ']
        for i in range(0,len(float_row_list)):
            for j in range(2,col_max):
                c = indexsummary.cell(row=float_row_list[i],column=j)
                c.value = float('%.2f'%kwargs[data_name[i]][j-2])
                c.font = datafont
        
        # 填入同比数据
        data_name = ['volumn_year_on_year',
        'year_on_year','east_year_on_year','mid_year_on_year','west_year_on_year',
        'year_on_year_under_90','year_on_year_90_144','year_on_year_above_144']
        for i in range(0,len(year_row_list)):
            for j in range(2,14):
                c = indexsummary.cell(row=year_row_list[i],column=j)
                c.value = '——'
                c.font = datafont
            for j in range(14,col_max):
                c = indexsummary.cell(row=year_row_list[i],column=j)
                c.value = '{:.2%}'.format(kwargs[data_name[i]][j-14])
                c.font = datafont
        
        # 填入环比数据
        data_name = ['volumn_chain',
        'chain','east_chain','mid_chain','west_chain',
        'chain_under_90','chain_90_144','chain_above_144']
        for i in range(0,len(year_row_list)):
            c = indexsummary.cell(row=chain_row_list[i],column=2)
            c.value = '——'
            c.font = datafont
            for j in range(3,col_max):
                c = indexsummary.cell(row=chain_row_list[i],column=j)
                c.value = '{:.2%}'.format(kwargs[data_name[i]][j-3])
                c.font = datafont
    
    def RadioPlot(self,**kwargs):
        # 参数列表:['index','year_on_year','chain','year_on_year_plot','chain_plot']

        # 表格基本格式设置
        radioplot = self.report['全国指数同比环比图']
        titlefont = Font(u'宋体',size = 12,bold=True,color='000000')
        datafont = Font(u'宋体',size = 10,bold=False,color='000000')
        bluefill = PatternFill("solid", fgColor='FFC0D9D9')
        rowmax = len(kwargs['index']) + 2

        # 填充颜色
        for i in range(1,5):
            c = radioplot.cell(row=1,column=i)
            c.fill = bluefill
        
        # 表格边框加粗
        for i in range(1,rowmax):
            for j in range(1,5):
                c = radioplot.cell(row=i,column=j)
                c.border = self.border
        
        # 填入表头
        c = radioplot.cell(row=1,column=2)
        c.value = '指数'
        c.font = titlefont
        c = radioplot.cell(row=1,column=3)
        c.value = '环比'
        c.font = titlefont
        c = radioplot.cell(row=1,column=4)
        c.value = '同比'
        c.font = titlefont

        # 填入日期
        for i in range(2,rowmax):
            c = radioplot.cell(row=i,column=1)
            c.value = self.to_date(i-2)
            c.font = titlefont
        
        # 填入数据
        for i in range(2,rowmax):
            c = radioplot.cell(row=i,column=2)
            c.value = float('%.2f'%kwargs['index'][i-2])
            c.font = datafont

            c = radioplot.cell(row=i,column=3)
            if i < 3:
                c.value = '——'
            else:
                c.value = float('%.2f'%(kwargs['chain'][i-3]*100))
            c.font = datafont

            c = radioplot.cell(row=i,column=4)
            if i < 14:
                c.value = '——'
            else:
                c.value = float('%.2f'%(kwargs['year_on_year'][i-14]*100))
            c.font = datafont
        
        # 添加图片
        img = Image(kwargs['year_on_year_plot'])
        img.width, img.height=1200, 300
        radioplot.add_image(img, 'E8')

        img = Image(kwargs['chain_plot'])
        img.width, img.height=1200, 300
        radioplot.add_image(img, 'E27')
    
    def SimplyPlot(self,**kwargs):
        # 参数列表:['url_index_plot','url_block_plot','url_line_plot']
        simplyplot = self.report['作图（简版）']
        simplyplot.column_dimensions['A'].height = 4.0
        titlefont = Font(u'宋体',size = 14,bold=True,color='000000')
        grayfill = PatternFill("solid", fgColor='00C0C0C0')

        # 填充颜色
        for j in [1,29,57]:
            for i in range(1,16):
                c = simplyplot.cell(row=j,column=i)
                c.fill = grayfill
        
        # 添加标题
        c = simplyplot.cell(row=1,column=1)
        c.value = '全国（简版封面&最简版图1）'
        c.font = titlefont
        c = simplyplot.cell(row=29,column=1)
        c.value = '各地区（最简版图2）'
        c.font = titlefont
        c = simplyplot.cell(row=57,column=1)
        c.value = '各面积（最简版图3）'
        c.font = titlefont

        # 添加图片
        img = Image(kwargs['url_index_plot'])
        img.width, img.height=1000, 500
        simplyplot.add_image(img, 'A2')
        img = Image(kwargs['url_block_plot'])
        img.width, img.height=1000, 500
        simplyplot.add_image(img, 'A30')
        img = Image(kwargs['url_line_plot'])
        img.width, img.height=1000, 500
        simplyplot.add_image(img, 'A58')

        # 合并单元格
        simplyplot.merge_cells('A1:F1')
        simplyplot.merge_cells('A29:F29')
        simplyplot.merge_cells('A57:F57')
    def CityVolumn(self,citylist :list,**kwargs):
        # 参数列表citylist中的城市的数据

        # 表格基本格式设置
        cityvolumn = self.report['各城市样本量']
        cityvolumn.freeze_panes = 'D3'
        cityvolumn.column_dimensions['A'].width = 5.0
        datafont = Font(u'宋体',size = 10,bold=False,color='000000')
        col_max = len(kwargs[citylist[0]]) + 3
        row_max = len(citylist) + 2

        # 填入表头时间
        for i in range(3,col_max):
            c = cityvolumn.cell(row=1,column=i)
            c.value = self.to_date(i-3)
            c.font = datafont
        # 填入序号
        for i in range(2,row_max):
            c = cityvolumn.cell(row=i,column=1)
            c.value = i-1
            c.font = datafont
            c = cityvolumn.cell(row=i,column=2)
            c.value = citylist[i-2]
            c.font = datafont
        # 填入数据
        for i in range(2,row_max):
            for j in range(3,col_max):
                c = cityvolumn.cell(row=i,column=j)
                c.value = kwargs[citylist[i-2]][j-3]
                c.font = datafont
    
    def CityIndex(self,citylist :list,**kwargs):
        # 参数列表citylist中的城市的数据

        # 表格基本格式设置
        cityindex = self.report['各城市指数']
        cityindex.freeze_panes = 'D3'
        cityindex.column_dimensions['A'].width = 5.0
        datafont = Font(u'宋体',size = 10,bold=False,color='000000')
        col_max = len(kwargs[citylist[0]]) + 3
        row_max = len(citylist) + 2

        # 填入表头时间
        for i in range(3,col_max):
            c = cityindex.cell(row=1,column=i)
            c.value = self.to_date(i-3)
            c.font = datafont
        # 填入序号
        for i in range(2,row_max):
            c = cityindex.cell(row=i,column=1)
            c.value = i-1
            c.font = datafont
            c = cityindex.cell(row=i,column=2)
            c.value = citylist[i-2]
            c.font = datafont
        # 填入数据
        for i in range(2,row_max):
            for j in range(3,col_max):
                c = cityindex.cell(row=i,column=j)
                c.value = float('%.2f'%kwargs[citylist[i-2]][j-3])
                c.font = datafont
    
    def CityPlot(self,citylist:list,poltlist:list):
        cityindex = self.report['各城市作图']
        datafont = Font(u'宋体',size = 10,bold=False,color='000000')
        
        # 25
        for i in range(0,len(citylist)):
            c = cityindex.cell(row=25*i + 1,column=1)
            c.value = '图' + str(i+1) + ' ' + citylist[i] + '\"城房指数\"'
            c.font = datafont
            img = Image(poltlist[i])
            img.width, img.height=1600, 400
            cityindex.add_image(img, 'A'+str(25*i+2))
    
    def CitySummary(self,informationlist:list):
        # list每项的参数列表:{'city_name','index_this_month','index_last_month','index_last_year',
        # 'chain_radio','year_on_year','volumn_chain','volumn_year'}
        
        # 表格基本格式设置
        citysummary = self.report['城市汇总']
        citysummary.freeze_panes = 'C2'
        citysummary.column_dimensions['A'].width = 5.0
        datafont = Font(u'宋体',size = 10,bold=False,color='000000')
        col_max = 10
        row_max = len(informationlist) + 2

        # 表格边框加粗
        for i in range(1,row_max):
            for j in range(1,col_max):
                c = citysummary.cell(row=i,column=j)
                c.border = self.border
        
        # 填入表头
        col_name_list = ['当月指数值','上月指数值','去年同月指数值',
        '环比','同比','交易量环比','交易量同比']
        for i in range(3,col_max):
            c = citysummary.cell(row=1,column=i)
            c.value = col_name_list[i-3]
            c.font = datafont
        
        # 填入数据
        col_data_list = ['city_name','index_this_month','index_last_month','index_last_year',
        'chain_radio','year_on_year','volumn_chain','volumn_year']
        for i in range(2,row_max):
            c = citysummary.cell(row=i,column=1)
            c.value = i-1
            c.font = datafont
            for j in range(2,col_max):
                c = citysummary.cell(row=i,column=j)
                if j == 2:
                    c.value = informationlist[i-2]['city_name']
                elif j < 6:
                    c.value = float('%.2f'%informationlist[i-2][col_data_list[j-2]])
                else:
                    c.value = '%.2f'%(informationlist[i-2][col_data_list[j-2]]*100) + '%'
                c.font = datafont

    def ComplexPlot(self,**kwargs):
        # 参数列表:['url_index_plot','url_block_plot','url_line_plot']
        complexplot = self.report['作图（详细）']

        # 添加图片
        img = Image(kwargs['url_index_plot'])
        img.width, img.height=1600, 400
        complexplot.add_image(img, 'A1')
        img = Image(kwargs['url_block_plot'])
        img.width, img.height=1600, 400
        complexplot.add_image(img, 'A26')
        img = Image(kwargs['url_line_plot'])
        img.width, img.height=1600, 400
        complexplot.add_image(img, 'A51')
    
    def CitySort(self,month:int,cityinformation:list):
        # 列表中每个元素应有参数:{'city_name','chain0','chain1' ,'chain2',
        # 'chain3' 'year0' ,'year1' ,'year2','year3'}

        # 表格基本格式设置
        citysort = self.report['城市排序']
        datafont = Font(u'宋体',size = 10,bold=True,color='000000')
        Reddatafont = Font(u'宋体',size = 10,bold=False,color='FF0000')
        bluefill = PatternFill("solid", fgColor='FF3299CC')
        rowmax = len(cityinformation) + 3
        colmax = 11
        align = Alignment(horizontal='center', vertical='center')

        # 表格边框加粗
        for i in range(1,rowmax):
            for j in range(1,colmax):
                c = citysort.cell(row=i,column=j)
                c.border = self.border

        # 颜色填充
        for j in range(1,colmax):
            for i in [1,2]:
                c = citysort.cell(row=i,column=j)
                c.fill = bluefill
        
        # 添加表头
        c = citysort.cell(row=1,column=3)
        c.value = '环比'
        c.font = datafont
        c.alignment = align
        citysort.merge_cells('C1:F1')

        c = citysort.cell(row=1,column=7)
        c.value = '同比'
        c.font = datafont
        c.alignment = align
        citysort.merge_cells('G1:J1')

        c = citysort.cell(row=2,column=1)
        c.value = '编号'
        c.font = datafont

        c = citysort.cell(row=2,column=2)
        c.value = '城市'
        c.font = datafont

        for i in range(0,4):
            c = citysort.cell(row=2,column=3+i)
            if month-i > 0:
                c.value = str(month-i) + '月'
            else:
                c.value = str(12+month-i) + '月'
            c.font = datafont
            c = citysort.cell(row=2,column=7+i)
            if month-i > 0:
                c.value = str(month-i) + '月'
            else:
                c.value = str(12+month-i) + '月'
            c.font = datafont
        
        # 填入数据
        col_data_list = ['city_name','chain0','chain1' ,'chain2','chain3' ,'year0' ,'year1' ,'year2','year3']
        for i in range(3,rowmax):
            c = citysort.cell(row=i,column=1)
            c.value = i-2
            c.font = datafont
            for j in range(2,colmax):
                c = citysort.cell(row=i,column=j)
                if j == 2:
                    c.value = cityinformation[i-3]['city_name']
                else:
                    c.value = float('%.2f'%(cityinformation[i-3][col_data_list[j-2]]*100))
                    if c.value == 0 or c.value == -100:
                        c.value = ''
                    elif c.value < 0:
                        c.font = Reddatafont

        
        # 添加排序功能
        citysort.auto_filter.ref = "A2:J" + str(colmax-1)
        citysort.auto_filter.add_sort_condition("A3:A"+ str(colmax -1))
