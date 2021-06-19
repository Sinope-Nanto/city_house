from worker.celery_app import app
from init_data.domains import init_database, init_city_index, init_total_data, init_city_list, \
    init_area_index, init_base_price_06, init_base_price_09, init_city_complex_info

@app.task
def initsystem(year, month):

    init_city_list()
    init_database(year,month)
    init_city_index()
    init_total_data()
    init_area_index()
    init_base_price_06()
    init_base_price_09()
    init_city_complex_info()

    return "初始化完成"