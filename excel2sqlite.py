#!/usr/bin/python3.7

import os
import sys
import logging
import pysqlite3 as sqlite3
import diagnostics as diag
from openpyxl import load_workbook


__title__ = sys.argv[0].split('/')[-1]
__author__ = "Fallflower"
__copyright__ = "Copyright 2020 lololol"
__credits__ = ["Fallflower"]
__license__ = "GPL"
__version__ = "0.5"
__maintainer__ = "Fallflower"
__email__ = "pitft@protonmail.com"
__status__ = "Production"


# if no log folder, create one.
if not os.path.exists('logs/'):
    os.mkdir('logs')

# initialize logger
logging.basicConfig(filename='logs/db_log.log', level=logging.NOTSET,
                    format=':%(asctime)s-%(levelname)s-%(lineno)s-%(message)s')


######
## OPENPYXL
## GRAB INFORMATION FROM EXCEL FILE

# Grab worksheet data
def grab_worksheet(xcel_filename, wkst_name):
    wb = load_workbook(filename=xcel_filename)
    return wb[wkst_name]


# access individual cells by --> wkst.cell(row,col)
# first_row --> wk.cell(1,1) == 'A1', wk.cell(2,1) == 'A2'
# first_col --> wk.cell(1,1) == 'A1', wk.cell(1,2) == 'B1'
def grab_fields(wkst, last_col, first_col=1, row=1, skip=[None, ' ']):
    fields = []
    types=[]
    for col in range(first_col, last_col):
        if wkst.cell(row,col).value in skip:
            continue

        fields.append(wkst.cell(row,col).value)
    return fields


# grab datatypes / data formats
def grab_types(wkst, last_col, first_col=1, row=1, skip=[None, ' ']):
    types = []
    for col in range(first_col, last_col):
        if wkst.cell(row,col).value in skip:
            continue
        types.append(wkst.cell(row,col).number_format)
    return types


# combine fields into a schema string for an sqlite statement
def create_schema(fields, types):
    schema=[]
    for num in range(len(fields)):
        if '.' in types[num]:
            types[num] = 'DOUBLE'
        elif '0' in types[num]:
            types[num] = 'INTEGER'
        elif types[num] == 'General':
            types[num] = 'TEXT'
        else:
            print('Unidentified number_format: ' + str(fields[num]) + " " + str(types[num]))
        schema.append(fields[num] + ' ' + types[num])

    return ", ".join(schema)


def grab_records(wkst, last_col, last_row, first_col=1, first_row=1, skip=[]):
    records = {}
    #starting at row 1 to the nth row.
    for row_num in range(first_row, last_row + 1):

        # if row is skippable, skip
        # Blank cell (row)                           or      'Item #' (2nd col)
        if wkst.cell(row_num,1).value in [None, ' '] or wkst.cell(row_num,2).value == "ItemNum":
            continue
        else:
            # else, prep the next row into the dictionary
            records.update({str(row_num):[]})

    # starting at 'A' (first col) to the nth col.
        for col_num in range(first_col, last_col):
            # if cell is skippable, skip.
            if col_num in skip:
                continue

            #if number formatted cell is empty, fill with 0
            if wkst.cell(row_num,col_num).value == None:
                wkst.cell(row_num,col_num).value = 0
            # append cell's value to list in the dictionary
                records[str(row_num)].append(stringify(wkst.cell(row_num,col_num).value))
    return records



def grab_records_gen_kinda(wkst, last_col, last_row, first_col=1, first_row=1, skip=[]):
    records = {}
    #starting at row 1 to the nth row.
    for row_num in range(first_row, last_row + 1):

        # if row is skippable, skip
        # Blank cell (row)                           or      'Item #' (2nd col)
        if wkst.cell(row_num,1).value in [None, ' '] or wkst.cell(row_num,2).value == "ItemNum":
            continue
        else:
            # else, prep the next row into the dictionary
            records.update({str(row_num):[]})

    # starting at 'A' (first col) to the nth col.
        for col_num in range(first_col, last_col):
            # if cell is skippable, skip.
            if col_num in skip:
                continue

            #if number formatted cell is empty, fill with 0
            if wkst.cell(row_num,col_num).value == None:
                wkst.cell(row_num,col_num).value = 0
            # append cell's value to list in the dictionary
            records[str(row_num)].append(stringify(wkst.cell(row_num,col_num).value))

        yield ', '.join(records[str(row_num)])
        records.clear()


def grab_letter(col_number):
    alphabet = {1:'A',2:'B',3:'C',4:'D',5:'E',6:'F',7:'G',8:'H',
                9:'I',10:'J',11:'K',12:'L',13:'M',14:'N',15:'O',
                16:'P',17:'Q',18:'R',19:'S',20:'T',21:'U',22:'V',
                23:'W',24:'X',25:'Y',26:'Z'}
    remainder = (col_number+1) % 26
    return alphabet[remainder]


def table_exists(cursor, table):
    cursor.execute(f'''SELECT count(name) FROM sqlite_master
                    WHERE type='table' AND name='{table}' ''')

    # If the count is 1, then table exists
    if cursor.fetchone()[0] == 1:
        logging.debug(f"'{table}' exists.")
        return True
    else:
        logging.debug(f"'{table}' doesn't exist.")
        return False


# Create a table.
def create_table(cursor, table, schema):
    # Check to see if table exists
    if table_exists(cursor, table):
        logging.debug(f"Table '{table}' already exists.")
        return  # breaking out of function

    # Schema - The field names and types of data the table is storing.
    cursor.execute(f'''CREATE TABLE {table} ({schema})''')
    logging.info(f"Table '{table} ({schema})' created.")


# list of fields --> string schema
#def create_schema(fields):
#    return ', '.join(fields)


# Create schema for table (manual)

# _m = manual
# _a = auto
def get_schema_m(schema={}):
    pass


# assuming filename has no '.' other than extension if at all
# if filename doesn't have ext, add ext.
def check_extension(filename, ext='.db'):
    if filename[-(len(ext)):] != ext:
        filename += ext
    return filename


# connects to database. returns connection object
# if database doesn't exist, sqlite3 creates it.
# So this will probably be the usual way to make
# new databases.
def connect(sqlite3_mod, database):
    database = check_extension(database)
    return sqlite3_mod.connect(database)


# Have to use the cursor object, so that means
# you must already have an sqlite connection to use this.
# creates database, doesn't return anything.
def create_db(cursor, database):
    database = check_extension(database)
    cursor.execute(f'{database}')

    # note: I guss I could do this by using subprocess module to
    #       to create the process of using sqlite3 cli tool to
    #       create the database then kill process (if it doesn't
    #       do it itself), but I don't feel like doing that.

# Debating keeping this.
def connect_to_db(sqlite3_mod, database):
    database = check_extension(database)
    connection = connect(sqlite3_mod, database)
    cursor = grab_cursor(connection)
    return connection, cursor


def grab_cursor(connection):
    return connection.cursor()


def commit(connection):
    connection.commit()


# close connection
def close(connection):
    connection.close()


# inserting data into some (or all, if you want) columns
def insert_cols(cursor, table, col_names=[], row_data=[], *args, **kwargs):
    cursor.execute(f'''INSERT INTO {table} ({",".join(col_names)}) VALUES ({','.join(row_data)})''')


# inserting data to every column
def insert_all(cursor, table, row_data):
    cursor.execute(f'''INSERT INTO {table} VALUES ({row_data})''')


# Prepping select queries for sending sqlite database into 2010` excel files
# select from certain columns
def select(cursor, table, cols):
    cursor.execute(f'''SELECT {','.join(cols)} from {table}''')

# select all columns
def select_all(cursor, table):
    cursor.execute(f'''SELECT * from {table}''')
    

def get_datatype(num):
    datatype = {'0': 'NULL', '1': 'INTEGER', '2': 'REAL',
                '3': 'TEXT', '4': 'NUMERIC', '5': 'BLOB'}

    return datatype[num] if num.isdigit() else False


def get_fields(cursor):
    pass


# sqlite really wants single quotes around text data.
# ex) s1 = 'string', s2 = '\'string\''
# cursor.execute(f"insert into table_a values ({s1})"
#   -> "insert into table_a values (string) x
#   -> Error: No column named string
# cursor.execute(f"insert into table_a values ({s2})"
#   -> "insert into table_a values ('string') ✔
def stringify(string):
    return (f'\'{str(string)}\'')

if __name__ == '__main__':
    # Basic program information
#    diag.__info__(__title__, __version__, __status__, __author__, __email__)
#
#    wb = grab_workbook()
    wk = grab_worksheet("dealer_price_list.xlsx", 'MI')

    # Need to add 1 to last_col since wkst.cell(row, col) starts at 1 for col instead of 0.
    last_col = len(wk[1]) + 1
    # len(wk[1]) --> amount of cells on row 1, starting at column 1 ('A')
    # to the last cell that has a value.

    last_row = len(wk['A'])
    # len(wk['A']) --> amount of cells on column 'A', starting at row 1
    # to the last cell that has a value.

    fs = grab_fields(wk, last_col, skip=[None,' ','Price Break'])
    ty = grab_types(wk, last_col, skip=[None,' ','Price Break'])

    try:
        database = sys.argv[1]
    except IndexError as e:
        # if no database (or file) is passed,
        # ask for one.
        database = input("Database: >")
        #if no database name, add one:
        if database == '\n':
            database = 'test.db'
	# if no .db extension, add it.
        database = check_extension(database)

    # Start connection to database
    logging.info("Attempting connection to \'{}\'.".format(database))

    # Connect to (or create) database.
    if not os.path.exists(database):
        logging.debug("{} doesn't exist. Creating database.".format(database))

    from datetime import datetime as dt

    try:
        # set connection and grab cursor
        t02=dt.now()
        connection, cursor = connect_to_db(sqlite3, database)
        t03=dt.now()
    except Exception as e:
          # incase anything happens
        logging.exception(e)
        raise e


    # Connection established
    logging.info('Connection established.')

    # Ask user what they want to do?

    # Get name for table
    table_name = input('table name:\n>')

    # Create table
    create_table(cursor, table_name, create_schema(fs, ty))

    from datetime import datetime

    # Insert some data
    t0=dt.now()
    for values in grab_records_gen_kinda(wk, last_col, last_row, skip=[None,' ',6]):
       insert_all(cursor, table_name, values)
    commit(connection)
    t1= dt.now()
    t2=dt.now()
    for k, recods in grab_records(wk, last_col, last_row, skip=[None,' ',6]).items():
        insert_all(cursor, table_name, ', '.join(recods))
    commit(connection)
    t3= dt.now()

    print('-' *50)
    print(f'gen: {(t1-t0).total_seconds()}')
    print('-' *50)
    print(f'grab: {(t3-t2).total_seconds()}')

#    print((t2-t0).total_seconds())
#    print((t2-t1).total_seconds())

#    print('-' *50)
#    print((t01-t00).total_seconds())
#    print((t03-t02).total_seconds())
#    print(((t03-t02)-(t01-t00)).total_seconds())

