import silsigan
from stock_list import import_list
from stock_day_datareader_pyun import stock_day_collector

import_list = import_list()
lists = import_list.run()

sdc = stock_day_collector()

sdc.run(lists)


