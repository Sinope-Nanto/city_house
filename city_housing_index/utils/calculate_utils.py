import xlrd
import numpy

from scipy import stats
from sklearn import datasets
from sklearn import linear_model

argc = 2
url = ['D:\\Code\\城房指数新编制\\1 指数编制2018-2020.7excel表\\2018-2020.7excel表\\33 成都标准数据2018.1.xls',
       'D:\\Code\\城房指数新编制\\1 指数编制2018-2020.7excel表\\2018-2020.7excel表\\33 成都标准数据2018.2.xls']


def getSB(data):  # 函数输入为excel文件的转置(二维数组，不含表头)，输出为标准住房(字典格式)
    SB = {
        "pro_id": stats.mode(data[1])[0][0],
        "unit_onsale": stats.mode(data[2])[0][0],
        "unit_duration": stats.mode(data[5])[0][0],
        "pro_area": numpy.mean(data[6]),
        "pro_floor": stats.mode(data[7])[0][0],
        "unit_floor": stats.mode(data[8])[0][0],
        "unit_area": numpy.mean(data[9]),
        "unit_price": numpy.mean(data[10]),
        "pro_dis": stats.mode(data[11])[0][0],
        "pro_block": stats.mode(data[12])[0][0],
        "block_ehn": stats.mode(data[13])[0][0],
        "block_edf": stats.mode(data[14])[0][0],
        "block_edn": stats.mode(data[15])[0][0],
        "block_enf": stats.mode(data[16])[0][0],
        "block_exn": stats.mode(data[17])[0][0],
        "block_exf": stats.mode(data[18])[0][0],
        "block_exb": stats.mode(data[19])[0][0],
        "block_ebf": stats.mode(data[20])[0][0],
        "block_edb": stats.mode(data[21])[0][0],
        "block_sdf": stats.mode(data[22])[0][0],
        "block_sdn": stats.mode(data[23])[0][0],
        "block_snf": stats.mode(data[24])[0][0],
        "block_sxn": stats.mode(data[25])[0][0],
        "block_sxf": stats.mode(data[26])[0][0],
        "block_sxb": stats.mode(data[27])[0][0],
        "block_sbf": stats.mode(data[28])[0][0],
        "block_sdb": stats.mode(data[29])[0][0],
        "block_rnf": stats.mode(data[30])[0][0],
        "block_rxf": stats.mode(data[31])[0][0],
        "subm_floor": stats.mode(data[32])[0][0],
        "subm_area": stats.mode(data[33])[0][0],
        "subm_putong": stats.mode(data[34])[0][0]}
    return SB


def get_ratio(data, data_lastmonth, data_lastyear):
    # 函数输入为三个二维数组，为excel的转置，不含表头，内容分别是：当期数据，上期数据，去年同期数据；输出为一个字典 year on year 为同比 chain为环比
    price = numpy.mean(data[10])
    return {"year_on_year": price / numpy.mean(data_lastyear[10]), "chain": price / numpy.mean(data_lastmonth[10])}


def linearRegression(data):  # 函数输入为excel文件的转置(二维数组，不含表头)，输出为回归方程结果(字典格式)
    reg = linear_model.LinearRegression()  # 声明模型
    num = len(data[0])  # 获得数据个数
    switch = num / 10  # 是否设置哑变量的阈值

    # 以下代码为添加哑变量
    table = []
    table_len = 0
    pro_id_list = []
    pro_id_number = []
    for i in range(0, num):
        newid = True
        for j in range(0, table_len):
            if data[1][i] == table[j]:
                newid = False
                pro_id_list[j].append(i)
                pro_id_number[j] += 1
                break
        if newid:
            table_len += 1
            table.append(data[1][i])
            pro_id_number.append(1)
            pro_id_list.append([i])
    no = 0
    dummy = []
    name_dummy = []
    for i in range(0, table_len):
        if pro_id_number[i] > switch:
            name_dummy.append(table[i])
            dummy.append([0 for k in range(0, num)])
            for j in range(0, pro_id_number[i]):
                dummy[no][pro_id_list[i][j]] = 1
            no += 1

    # 向模型中装入数据
    dataset = []
    for i in range(0, num):
        sample = [data[2][i]]
        sample += [data[j][i] for j in range(5, 10)]
        sample += [data[j][i] for j in range(11, 35)]
        for j in range(0, no):
            sample.append(dummy[j][i])
        dataset.append(sample)

        # 开始拟合
    reg.fit(X=dataset, y=data[10])

    # 返回结果
    result = {
        "intercept": reg.intercept_,
        "unit_onsale": reg.coef_[0],
        "unit_duration": reg.coef_[1],
        "pro_area": reg.coef_[2],
        "pro_floor": reg.coef_[3],
        "unit_floor": reg.coef_[4],
        "unit_area": reg.coef_[5],
        "pro_dis": reg.coef_[6],
        "pro_block": reg.coef_[7],
        "block_ehn": reg.coef_[8],
        "block_edf": reg.coef_[9],
        "block_edn": reg.coef_[10],
        "block_enf": reg.coef_[11],
        "block_exn": reg.coef_[12],
        "block_exf": reg.coef_[13],
        "block_exb": reg.coef_[14],
        "block_ebf": reg.coef_[15],
        "block_edb": reg.coef_[16],
        "block_sdf": reg.coef_[17],
        "block_sdn": reg.coef_[18],
        "block_snf": reg.coef_[19],
        "block_sxn": reg.coef_[20],
        "block_sxf": reg.coef_[21],
        "block_sxb": reg.coef_[22],
        "block_sbf": reg.coef_[23],
        "block_sdb": reg.coef_[24],
        "block_rnf": reg.coef_[25],
        "block_rxf": reg.coef_[26],
        "subm_floor": reg.coef_[27],
        "subm_area": reg.coef_[28],
        "subm_putong": reg.coef_[29]}
    for i in range(0, no):
        result["dummy_pro_id" + str(name_dummy[i])] = reg.coef_[30 + i]
    return result


def _DataRead(argc, urls):
    text = xlrd.open_workbook(urls[0])
    worksheet = text.sheet_by_index(0)
    ncols = worksheet.ncols
    data = [[] for i in range(0, 35)]
    for i in range(0, argc):
        text = xlrd.open_workbook(urls[i])
        worksheet = text.sheet_by_index(0)
        for j in range(0, ncols):
            data[j] += worksheet.col_values(j)[1:]
    return data
