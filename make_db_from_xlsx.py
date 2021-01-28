import excel2sqlite as e2s
from datetime import datetime as dt
from diagnostics import test_time
xcel_file = 'pv_price_list.xlsx'
wkst_name = 'MI'
database = 'Peavey'
tb ='MI'


wkst = e2s.grab_worksheet(xcel_file, wkst_name)

connection, cursor = connect_to_db(wkst, database)


