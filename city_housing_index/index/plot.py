import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import numpy as np

class PlotReport:
    def to_date_year_on_year_index(self, temp, position):
        year = int(temp)//12 + 6
        month = int(temp)%12 + 1
        return str(year).zfill(2) + '.' + str(month).zfill(2)  + '     '
    def to_date_chain(self, temp, position):
        year = int(temp)//12 + 6
        month = int(temp)%12 + 1
        return str(year).zfill(2) + '.' + str(month).zfill(2)  + '                               '
    def to_1_10k(self, temp, position):
        return str('%.1f'%(temp/10000))
    def to_percent(self, temp, position):
        return '%.0f'%(100*temp) + '%'
    def to_date_index_vol(self, temp, position):
        year = int(temp)//12 + 2006
        month = int(temp)%12 + 1
        if month == 1:
            return str(year) + '-' + str(month)
        else:
            return str(month)
    def plot_year_on_year_radio(self,data,saveurl):
        #全国同比图
        fig, ax = plt.subplots(figsize=(20, 5))  
        ax.spines['bottom'].set_position(('data',0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        ax.spines['left'].set_position(('data',11))  #data表示通过值来设置y轴的位置
        # 取消边框
        for key, spine in ax.spines.items():
            # 'left', 'right', 'bottom', 'top'
            if key == 'right' or key == 'top':
                spine.set_visible(False)
        plt.plot(range(12,len(data) + 12) , data , color='black',marker='.'
            ,markeredgecolor='black',markersize='5',linewidth =1)
        plt.xlim(xmin = 11)
        plt.gca().yaxis.set_major_formatter(FuncFormatter(self.to_percent))
        plt.gca().xaxis.set_major_formatter(FuncFormatter(self.to_date_year_on_year_index))
        plt.xticks(range(12,len(data) + 12,4)) 
        plt.xticks(rotation=90, fontsize=8)
        plt.savefig(saveurl,dpi=200)
        plt.close()   
    def plot_chain_radio(self,data,saveurl):
        #全国环比图
        fig, ax = plt.subplots(figsize=(20, 5))  
        ax.spines['bottom'].set_position(('data',0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        ax.spines['left'].set_position(('data',1))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        # 取消边框
        for key, spine in ax.spines.items():
            # 'left', 'right', 'bottom', 'top'
            if key == 'right' or key == 'top':
                spine.set_visible(False)
        plt.plot(range(1,len(data) + 1) , data , color='black',marker='.'
            ,markeredgecolor='black',markersize='5',linewidth =1)
        plt.xlim(xmin = 1)
        plt.gca().yaxis.set_major_formatter(FuncFormatter(self.to_percent))
        plt.gca().xaxis.set_major_formatter(FuncFormatter(self.to_date_chain))
        plt.xticks(range(1,len(data) + 1,4)) 
        plt.xticks(rotation=90, fontsize=8)
        plt.savefig(saveurl,dpi=200)
        plt.close()
    def plot_vol_index(self,volume,index,saveurl):
        #全国指数-销量图

        fig, ax = plt.subplots(figsize=(20, 5))
        ax.bar(range(0,len(volume)),volume,0.3,color='white',ec = 'gray',lw = 1,hatch = '-')
        ax.spines['left'].set_position(('data',-1))
        ax.set_xlim(xmin = -1)
        ax.yaxis.set_major_formatter(FuncFormatter(self.to_1_10k))
        ax.xaxis.set_major_formatter(FuncFormatter(self.to_date_index_vol))
        ax.set_ylabel('交易量/万套',fontproperties = 'SimHei')
        ax.set_xticks(range(0,len(volume),4))
        plt.xticks(rotation=90, fontsize=8)
        ax2 = ax.twinx()
        ax2.plot(range(0,len(volume)), index, linewidth=1,color='black',marker = '.',
        markeredgecolor='black',markersize='3')
        ax2.set_yticks(range(100,int(max(index)+100),100))
        ax2.set_ylabel('指数',fontproperties = 'SimHei')
        plt.savefig(saveurl,dpi=200)
        plt.close()
    def plot_index(self,data,saveurl):
        # 全国指数图
        fig, ax = plt.subplots(figsize=(10, 5))  
        ax.spines['bottom'].set_position(('data',100))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        ax.spines['left'].set_position(('data',0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        # 取消边框
        for key, spine in ax.spines.items():
            # 'left', 'right', 'bottom', 'top'
            if key == 'right' or key == 'top':
                spine.set_visible(False)
        plt.plot(range(0,len(data)) , data , color='black',marker='.'
            ,markeredgecolor='black',markersize='2',linewidth =1)
        plt.gca().xaxis.set_major_formatter(FuncFormatter(self.to_date_index_vol))
        plt.xlim(xmin = 0)
        plt.ylim(ymin = 100)
        plt.xticks(range(0,len(data),4))
        plt.yticks(range(100,int(max(data)+50),50))
        plt.grid(axis="y",linestyle='--',color='gray')
        plt.xticks(rotation=90, fontsize=6)
        plt.yticks(fontsize=6)
        plt.savefig(saveurl,dpi=200)
        plt.close()

    def plot_index_by_block(self,data,saveurl):

        plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
        plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
        # 全国地区指数图
        fig, ax = plt.subplots(figsize=(10, 5))  
        ax.spines['bottom'].set_position(('data',100))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        ax.spines['left'].set_position(('data',0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        # 取消边框
        for key, spine in ax.spines.items():
            # 'left', 'right', 'bottom', 'top'
            if key == 'right' or key == 'top':
                spine.set_visible(False)
        plt.plot(range(0,len(data[0])) , data[0] , color='black',marker='.'
            ,markeredgecolor='black',markersize='2',linewidth =1,label='东部指数')
        plt.plot(range(0,len(data[1])) , data[1] , color='blue',marker='.'
            ,markeredgecolor='blue',markersize='2',linewidth =1,label='中部指数')
        plt.plot(range(0,len(data[2])) , data[2] , color='red',marker='.'
            ,markeredgecolor='red',markersize='2',linewidth =1,label='西部指数')
        plt.legend() # 显示图例
        plt.gca().xaxis.set_major_formatter(FuncFormatter(self.to_date_index_vol))
        plt.xlim(xmin = 0)
        plt.ylim(ymin = 100)
        plt.xticks(range(0,len(data[0]),4))
        plt.yticks(range(100,int(max([max(data[0]),max(data[1]),max(data[2])])+50),50))
        plt.grid(axis="y",linestyle='--',color='gray')
        plt.xticks(rotation=90, fontsize=6)
        plt.yticks(fontsize=6)
        plt.savefig(saveurl,dpi=200)
        plt.close()

    def plot_index_by_area(self,data,saveurl):

        plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
        plt.rcParams['axes.unicode_minus']=False #用来正常显示负号

        # 全国地区指数图
        fig, ax = plt.subplots(figsize=(10, 5))  
        ax.spines['bottom'].set_position(('data',100))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        ax.spines['left'].set_position(('data',0))  #data表示通过值来设置x轴的位置，将x轴绑定在y=0的位置
        # 取消边框
        for key, spine in ax.spines.items():
            # 'left', 'right', 'bottom', 'top'
            if key == 'right' or key == 'top':
                spine.set_visible(False)
        plt.plot(range(0,len(data[0])) , data[0] , color='black',marker='.'
            ,markeredgecolor='black',markersize='2',linewidth =1,label='小户型指数')
        plt.plot(range(0,len(data[1])) , data[1] , color='blue',marker='.'
            ,markeredgecolor='blue',markersize='2',linewidth =1,label='中户型指数')
        plt.plot(range(0,len(data[2])) , data[2] , color='red',marker='.'
            ,markeredgecolor='red',markersize='2',linewidth =1,label='大户型指数')
        plt.legend() # 显示图例
        plt.gca().xaxis.set_major_formatter(FuncFormatter(self.to_date_index_vol))
        plt.xlim(xmin = 0)
        plt.ylim(ymin = 100)
        plt.xticks(range(0,len(data[0]),4))
        plt.yticks(range(100,int(max([max(data[0]),max(data[1]),max(data[2])])+50),50))
        plt.grid(axis="y",linestyle='--',color='gray')
        plt.xticks(rotation=90, fontsize=6)
        plt.yticks(fontsize=6)
        plt.savefig(saveurl,dpi=200)
        plt.close()    

    def plot_vol_index_by_block_complex(self,volume,index,saveurl):
        # 全国指数-销量图 -复杂版 -按地区分
        plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
        plt.rcParams['axes.unicode_minus']=False #用来正常显示负号

        fig, ax = plt.subplots(figsize=(20, 5))
        ax.bar(range(0,len(volume[0])),volume[0],0.20,color='white',ec = 'gray',lw = 1,hatch = '-',label='东部交易量')
        ax.bar(np.arange(len(volume[1]))+0.20,volume[1],0.20,color='white',ec = 'gray',lw = 1,hatch = '/',label='中部交易量')
        ax.bar(np.arange(len(volume[2]))+0.40,volume[2],0.20,color='white',ec = 'gray',lw = 1,hatch = '.',label='西部交易量')
        ax.spines['left'].set_position(('data',-1))
        ax.legend(loc=1)
        ax.set_xlim(xmin = -1)
        ax.yaxis.set_major_formatter(FuncFormatter(self.to_1_10k))
        ax.xaxis.set_major_formatter(FuncFormatter(self.to_date_index_vol))
        ax.set_ylabel('交易量/万套',fontproperties = 'SimHei')
        ax.set_xticks(range(0,len(volume[0]),4))
        plt.xticks(rotation=90, fontsize=8)
        ax2 = ax.twinx()
        ax2.plot(range(0,len(index[0])) , index[0] , color='black',marker='.'
            ,markeredgecolor='black',markersize='2',linewidth =1,label='东部指数')
        ax2.plot(range(0,len(index[1])) , index[1] , color='blue',marker='.'
            ,markeredgecolor='blue',markersize='2',linewidth =1,label='中部指数')
        ax2.plot(range(0,len(index[2])) , index[2] , color='red',marker='.'
            ,markeredgecolor='red',markersize='2',linewidth =1,label='西部指数')
        ax2.set_yticks(range(100,int(max([max(index[0]),max(index[1]),max(index[2])])+150),50))
        ax2.set_ylabel('指数',fontproperties = 'SimHei')
        ax2.legend(loc=2)
        plt.savefig(saveurl,dpi=200)
        plt.close()

    def plot_vol_index_by_area_complex(self,volume,index,saveurl):
        # 全国指数-销量图 -复杂版 -按面积分
        plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
        plt.rcParams['axes.unicode_minus']=False #用来正常显示负号

        fig, ax = plt.subplots(figsize=(20, 5))
        ax.bar(range(0,len(volume[0])),volume[0],0.20,color='white',ec = 'gray',lw = 1,hatch = '-',label='小户型交易量')
        ax.bar(np.arange(len(volume[1]))+0.20,volume[1],0.20,color='white',ec = 'gray',lw = 1,hatch = '/',label='中户型交易量')
        ax.bar(np.arange(len(volume[2]))+0.40,volume[2],0.20,color='white',ec = 'gray',lw = 1,hatch = '.',label='大户型交易量')
        ax.spines['left'].set_position(('data',-1))
        ax.legend(loc=1)
        ax.set_xlim(xmin = -1)
        ax.yaxis.set_major_formatter(FuncFormatter(self.to_1_10k))
        ax.xaxis.set_major_formatter(FuncFormatter(self.to_date_index_vol))
        ax.set_ylabel('交易量/万套',fontproperties = 'SimHei')
        ax.set_xticks(range(0,len(volume[0]),4))
        plt.xticks(rotation=90, fontsize=8)
        ax2 = ax.twinx()
        ax2.plot(range(0,len(index[0])) , index[0] , color='black',marker='.'
            ,markeredgecolor='black',markersize='2',linewidth =1,label='小户型指数')
        ax2.plot(range(0,len(index[1])) , index[1] , color='blue',marker='.'
            ,markeredgecolor='blue',markersize='2',linewidth =1,label='中户型指数')
        ax2.plot(range(0,len(index[2])) , index[2] , color='red',marker='.'
            ,markeredgecolor='red',markersize='2',linewidth =1,label='大户型指数')
        ax2.set_yticks(range(100,int(max([max(index[0]),max(index[1]),max(index[2])])+150),50))
        ax2.set_ylabel('指数',fontproperties = 'SimHei')
        ax2.legend(loc=2)
        plt.savefig(saveurl,dpi=200)
        plt.close()   