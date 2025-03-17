from sqlalchemy.engine import URL
from sqlalchemy import create_engine
import pandas as pd
import os
import pyodbc  # Needed for multiple result sets
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get data from .env
DRIVER_NAME = os.getenv("DRIVER_NAME")
DRIVER_NAME_SAMO = os.getenv("DRIVER_NAME_SAMO")
SERVER_NAME_RO = os.getenv("SERVER_NAME_RO")
DATABASE_NAME_RO = os.getenv("DATABASE_NAME_RO")  # Ensure consistency
SERVER_NAME_SAMO = os.getenv("SERVER_NAME_SAMO")
DATABASE_NAME_SAMO = os.getenv("DATABASE_NAME_SAMO")  # Fixed naming
PASSWORD = os.getenv("PASSWORD")
USERNAME = os.getenv("USER")
PASSWORD_SAMO = os.getenv("PASSWORD_SAMO")
USERNAME_SAMO = os.getenv("USER_SAMO")


class ConnectionHandler:
    def __init__(self):
        # Build the connection string
        driver = DRIVER_NAME
        driver_samo = DRIVER_NAME_SAMO
        connection_string = f"DRIVER={driver};SERVER={SERVER_NAME_RO};DATABASE={DATABASE_NAME_RO};UID={USERNAME};PWD={PASSWORD}"
        samo_con_string = f"DRIVER={driver};SERVER={SERVER_NAME_SAMO};DATABASE={DATABASE_NAME_SAMO};UID={USERNAME_SAMO};PWD={PASSWORD_SAMO}"
        
        # Create SQLAlchemy engines
        connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
        self.engine = create_engine(connection_url, use_setinputsizes=False, echo=False)
        self.raw_connection = pyodbc.connect(connection_string) 
        
        samo_con_url = URL.create("mssql+pyodbc", query={"odbc_connect": samo_con_string})
        self.engine_samo = create_engine(samo_con_url, use_setinputsizes=False, echo=False)
        self.raw_connection_samo = pyodbc.connect(samo_con_string)

    def fetch_data(self, query, db):
        """Fetch data using a SQL query."""
        try:
            if db == "RO":
                return pd.read_sql(query, self.engine)  # Use SQLAlchemy engine
            elif db == "SAMO":
                return pd.read_sql(query, self.engine_samo)
        except Exception as e:
            print(e)


    def execute_procedure(self, proc_name, params, db):
        """Call a stored procedure with parameters and return multiple result sets."""
        try:
            if db == "RO":
                conn = self.raw_connection
            elif db == "SAMO":
                conn = self.raw_connection_samo
            else:
                raise ValueError("Invalid database selection")

            cursor = conn.cursor()

            # Split params safely
            param_list = [p.strip() for p in params.split(';') if p.strip()]
            
            if not param_list:
                raise ValueError("No parameters provided for stored procedure execution")

            # Convert types if needed (assumes second param is integer)
            if len(param_list) > 1:
                print(repr(param_list[1]))

                param_list[1] = int(param_list[1])  # Convert second param to int
            print(type(param_list[1]))
            param_placeholders = ', '.join(['?' for _ in param_list])
            sql = f"EXEC {proc_name} {param_placeholders}"
        
            print(f"Executing: {sql} with params: {param_list}")


            cursor.execute(sql, tuple(param_list))  # Pass as tuple

            results = []
            while True:
                try:
                    while True:
                        if cursor.description:  # Only fetch if there are columns
                            columns = [col[0] for col in cursor.description]
                            rows = cursor.fetchall()
                            if rows:  # Ensure rows are fetched
                                df = pd.DataFrame.from_records(rows, columns=columns)
                                results.append(df)
                            else:
                                print("⚠️ Query executed but returned no rows.")
                        else:
                            print("⚠️ Cursor description is None. No data found.")

                        if not cursor.nextset():
                            break  # Move to the next result set
                    
                    
                    
                    print('test')
                    df = pd.DataFrame.from_records(cursor.fetchall(), columns=[col[0] for col in cursor.description])
                    print('test')

                    results.append(df)
                except pyodbc.ProgrammingError:
                    break  # No more result sets
                if not cursor.nextset():
                    break

            return results if len(results) > 1 else (results[0] if results else pd.DataFrame())

        except Exception as e:
            print(f"Error executing stored procedure: {e}")
            return None

    def close_connection(self, db):
        """Close the database connection."""
        if db == "RO":
            self.raw_connection.close()
            self.engine.dispose()
        elif db == "SAMO":
            self.raw_connection_samo.close()
            self.engine_samo.dispose()