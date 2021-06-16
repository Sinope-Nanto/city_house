import multiprocessing

bind = "0.0.0.0:3000"
workers = multiprocessing.cpu_count() * 2 + 1
timeout = 300
BASE_DIR = "/root/city_house/city_housing_index"
accesslog = '/var/log/gunicorn/city_index_access.log'
errorlog = '/var/log/gunicorn/city_index_error.log'
# access_log_format = '%(h)s %(l)s %(u)s %(t)s "%(r)s" %(s)s %(b)s "%(f)s" "%(a)s"'
# access_log_format = "%(t)s %(p)s %(h)s \"%(r)s\" %(s)s %(L)s %(b)s \"%(f)s\" \"%(a)s\""
access_log_format = '%(t)s %(p)s %(h)s "%(r)s" %(s)s %(L)s %(b)s "%(f)s"'
