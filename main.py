import sqlite3
import pandas as pd
import config


def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        print("Connection established")
    except sqlite3.Error as e:
        print(e)
    return conn


# Load data from Excel file into SQLite database
def load_data_to_sqlite(conn, excel_file):
    df = pd.read_excel(excel_file, engine='openpyxl')
    df.to_sql('online_retail', conn, if_exists='replace', index=False)
    print("Data loaded into SQLite database")


def main():
    database = config.DATABASE_PATH
    conn = create_connection(database)
    if conn is not None:
        load_data_to_sqlite(conn, 'Online retail.xlsx')
        conn.close()  # Close the connection
    else:
        print("Error! Cannot create the database connection.")


if __name__ == '__main__':
    main()