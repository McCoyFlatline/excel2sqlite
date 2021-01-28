import pysqlite3 as sqlite3
import logging


logging.basicConfig(filename='db_log.log', level=logging.NOTSET,
                    format=':%(asctime)s-%(levelname)s-%(lineno)s-%(message)s')


class E2S_Converter_Master(object):
    def __init__(self, database, xcel=[], wb=[]):
        self.__database = database
        self.__connection = sqlite3.connect(self.__database)
        self.__cursor = self.__connection.cursor()

    def create_table(self, table, schema_d):
        # Check to see if table exists
        if self.table_exists(table):
            logging.debug("Table '{}' already exists.".format(table))
            return  # breaking out of function

        # Schema - The field names and types of data the table is storing.
        schema = []
        for field, datatype in schema_d.items():
            schema.append(field + " " + datatype)

        self.__cursor.execute("""CREATE TABLE {}({})
            """.format(table, ",".join(schema)))
        logging.info("Table '{}' created.".format(table))

    # table_exists
    # Check if table exists
    def table_exists(self, table):
        self.__cursor.execute('''SELECT count(name) FROM sqlite_master
            WHERE type='table' AND name='{}' '''.format(table))

        # If the count is 1, then table exists
        if self.__cursor.fetchone()[0] == 1:
            return True
        else:
            return False

    # alter_table
    # Sqlite only really supports table & column renaming, and adding a column.
    # Alter table will remain an empty method until further notice.
    def alter_table(self, table):
        pass

    def create_schema(self, schema={}):
        fields_done = 0
        num_of_fields = int(input("How many fields (columns)? "))

        if num_of_fields <= 0:
            return

        field = ""
        data_type = ""
        while fields_done < num_of_fields:
            field = input("field Key Value: ")
            if field in ["no", "n", "leave", "back", "exit"]:
                return schema
            else:
                data_type = input("Choose which data type: ")
                if data_type in ["null", "integer", "text", "float", "real", "blob"]:
                    schema[field] = data_type.upper()
                elif int(data_type) >= 1 and int(data_type) <= 6:
                    schema[field] = self.get_datatype(int(data_type))
                else:
                    print("Redo Key and data type.")
                    fields_done -= 1
                    continue

                # if at last loop iteration, ask if any more keys to add.
                # if yes, extend the loop by subtracting the loop's counter.
                # In this case, it's 'fields_done'
                # Cool little optional feature imo.
                if fields_done == num_of_fields - 1:
                    field = input("Any more? ").lower()
                    if field in ["yes", "y"]:
                        fields_done -= 1
                    elif field.isdigit():
                        if int(field) >= 1:
                            fields_done -= int(field)
                    else:
                        return schema

                # increment loop counter
                fields_done += 1

    def insert(self):
        pass

    def drop_table(self):
        pass

    def commit(self):
        pass

    def add_column(self):
        pass

    def rename_column(self):
        pass

    def rename_table(self):
        pass

    def get_datatype(num):
        if num == 1:
            return 'NULL'
        elif num == 2:
            return 'INTEGER'
        elif num == 3:
            return 'TEXT'
        elif num == 4:
            return 'FLOAT'
        elif num == 5:
            return 'REAL'
        elif num == 6:
            return 'BLOB'
        else:
            print('No data_type')
            return

    # Thank you PinnyM on StackOverflow
    # /questions/16900552/change-the-primary-key-of-a-table-in-sqlite
    def add_primary_key(self):
        self.create_table('', {})
        self.insert()
        # deciding if to use alter_table with 'RENAME' parameter
        # or having a specific rename method
        self.alter_table()
        self.commit()

    def copy_table(self):
        pass

    # getters
    @property
    def database(self):
        return self.__database


class E2S(E2S_Converter_Master):
    pass


if __name__ == '__main__':
    db = 'test.db'
    converter = E2S_Converter_Master(db)
