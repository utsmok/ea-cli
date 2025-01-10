import json
import sqlite3

def store_list_as_table(db_conn, table_name, data_list):
    # Collect all possible keys
    all_keys = set()
    for item in data_list:
        all_keys.update(item.keys())

    # Create table columns
    columns = [f"{key} TEXT" for key in all_keys]
    create_sql = f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join(columns)})"
    db_conn.execute(create_sql)

    # Prepare insert statement
    placeholders = ", ".join(["?" for _ in all_keys])
    insert_sql = f"INSERT INTO {table_name} ({', '.join(all_keys)}) VALUES ({placeholders})"

    # Insert each item
    for item in data_list:
        row_values = []
        for key in all_keys:
            val = item.get(key)
            if isinstance(val, list):
                if len(val) == 0:
                    val = None
                elif len(val) == 1:
                    val = val[0]
                else:
                    val = json.dumps(val)
            if isinstance(val, dict):
                val = json.dumps(val)
            row_values.append(val)
        try:
            db_conn.execute(insert_sql, row_values)
        except Exception as e:
            print(row_values)
            raise e

    db_conn.commit()

with open("enriched_data.json", "r") as f:
    detailed_data_list = json.load(f)
with open("person_data.json", "r") as f:
    person_data_list = json.load(f)

conn = sqlite3.connect('database.db')
store_list_as_table(conn, "detailed_data", detailed_data_list)
store_list_as_table(conn, "detailed_person_data", person_data_list)
conn.close()