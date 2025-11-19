import pandas as pd
import mysql.connector

conn = mysql.connector.connect(
    host='localhost',
    user='znfu',
    password='laravel',
    database='laravel'
)

query = """
SELECT * FROM dfas
"""

# Read data into a pandas DataFrame
df = pd.read_sql(query, conn)

# Export to Excel
df.to_excel('dfas.xlsx', index=False)
print("File dfas.xlsx created successfully!")