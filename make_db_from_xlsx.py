import excel2sqlite as e2s
from datetime import datetime as dt
from diagnostics import test_time
xcel_file = 'pv_price_list.xlsx'
wkst_name = 'MI'
database = 'Peavey'
table ='MI'

#load excel filename and grab worksheet by name
wkst = e2s.grab_worksheet(xcel_file, wkst_name)

# Find position of last cell with value along 1st row and 'A' column
last_row = wkst[1]
last_col = wkst['A']

fields = e2s.grab_fields(wkst, last_col)
types = e2s.grab_types()
schema = e2s.create_schema()

connection, cursor = e2s.connect_to_db(database)


e2s.create_table(cursor, table, schema)

for k, record in e2s.grab_records(wkst, last_col,last_row,skip=skipp).items():
    e2s.insert_all(cursor, table, ','.join(record)

e2s.commit(connection)



def test(wkst=wkst, last_col=last_col, last_row=last_row, skipp=[None,' ',7]):
    r=[]
    t0 = dt.now()
    r = e2s.grab_records(wkst, last_col,last_row,skip=skipp).items()
    for k, record in r:
        e2s.insert_all(cursor, table, ','.join(record))
    e2s.commit(connection)
    t1 = dt.now()

    t2 = dt.now()
    for k, record in e2s.grab_records(wkst, last_col,last_row,skip=skipp).items():
        e2s.insert_all(cursor, table, ','.join(record)
    e2s.commit(connection)
    t3 = dt.now()

    t4 = dt.now()
    for record in e2s.grab_records_gen_kinda(wkst, last_col,last_row,skip=skipp):
        e2s.insert_all(cursor, table, record)
    e2s.commit(connection)
    t5 = dt.now()

    print(f'list first: {(t1-t0).total_seconds()}')
    print(f'grab: {(t3-t2).total_seconds()}')
    print(f'gen: {(t5-t4).total_seconds()}')
