import pandas as pd
import mysql.connector
import os

def get_script_directory():
    """Get the directory where the current script is located"""
    return os.path.dirname(os.path.abspath(__file__))

def create_db_connection():
    """Create a connection to the MySQL database"""
    try:
        conn = mysql.connector.connect(
            host='database.primenetpay.com',
            user='primenetadmin',
            password='45677',
            database='gated',
            port=3306
        )
        return conn
    except mysql.connector.Error as e:
        print(f"Error connecting to database: {e}")
        return None

def read_excel_file(file_path):
    """Read the Excel file"""
    try:
        if not os.path.exists(file_path):
            print(f"Excel file not found: {file_path}")
            return None
        
        df = pd.read_excel(file_path, header=None)
        df.columns = ['external_id', 'merchant_id']
        
        print(f"Found {len(df)} rows in Excel")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def update_transactions(conn, excel_data):
    """Update transactions in the database"""
    try:
        cursor = conn.cursor()
        update_query = """
        UPDATE transactions 
        SET merchant_id = %s, status_code = 300, response_message = 'SUCCESS'
        WHERE external_id = %s
        """
        
        updated_count = 0
        
        for idx, row in excel_data.iterrows():
            merchant_id = str(row['merchant_id'])
            external_id = str(row['external_id'])
            
            cursor.execute(update_query, (merchant_id, external_id))
            updated_count += cursor.rowcount
        
        conn.commit()
        cursor.close()
        
        print(f"Updated {updated_count} transactions")
        return updated_count
    except Exception as e:
        print(f"Error updating database: {e}")
        return 0

def main():
    script_dir = get_script_directory()
    excel_file_path = os.path.join(script_dir, 'transactions.xlsx')
    
    excel_data = read_excel_file(excel_file_path)
    if excel_data is None:
        return
    
    conn = create_db_connection()
    if conn is None:
        return
    
    update_transactions(conn, excel_data)
    
    conn.close()
    print("Done")

if __name__ == "__main__":
    main()