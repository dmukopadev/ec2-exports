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
            password='fr6OznzojH2mpkUgXQQQ',
            database='primenet_payment_gateway'
        )
        print("✅ Successfully connected to the database")
        return conn
    except mysql.connector.Error as e:
        print(f"❌ Error connecting to database: {e}")
        return None

def read_excel_file(file_path):
    """Read the Excel file with columns 'external_id' and 'order_id'"""
    try:
        if not os.path.exists(file_path):
            print(f"❌ Excel file not found: {file_path}")
            return None
        
        df = pd.read_excel(file_path, sheet_name='Transactions')
        print(f"✅ Successfully read Excel file: {file_path}")
        
        # Check required columns
        required_cols = ['external_id', 'order_id']
        if not all(col in df.columns for col in required_cols):
            print(f"❌ Excel file must contain columns: {required_cols}")
            return None
        
        print(f"📊 Found {len(df)} rows in Excel file")
        return df
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return None

def update_transactions(conn, excel_data):
    """Update transactions in the database"""
    try:
        cursor = conn.cursor()
        update_query = """
        UPDATE transactions
        SET account_number = %s
        WHERE external_id = %s
        """
        
        updated_count = 0
        
        for idx, row in excel_data.iterrows():
            order_id = str(row['order_id'])
            external_id = str(row['external_id'])
            
            cursor.execute(update_query, (order_id, external_id))
            updated_count += cursor.rowcount  # rowcount = 1 if row updated, 0 if no match
        
        conn.commit()
        cursor.close()
        
        print(f"✅ Database update complete. Total rows updated: {updated_count}")
        return updated_count
    except Exception as e:
        print(f"❌ Error updating database: {e}")
        return 0

def main():
    script_dir = get_script_directory()
    excel_file_path = os.path.join(script_dir, 'orders.xlsx')
    
    print("🚀 Starting transaction update process...")
    print(f"📁 Script directory: {script_dir}")
    print(f"📊 Input Excel: {excel_file_path}")
    
    # Step 1: Read Excel file
    excel_data = read_excel_file(excel_file_path)
    if excel_data is None:
        print("❌ Failed to read Excel file. Exiting.")
        return
    
    # Step 2: Connect to database
    conn = create_db_connection()
    if conn is None:
        print("❌ Failed to connect to database. Exiting.")
        return
    
    # Step 3: Update transactions
    updated_count = update_transactions(conn, excel_data)
    
    # Close database connection
    conn.close()
    print("✅ Database connection closed")
    
    print(f"🎉 Process completed! {updated_count} transactions updated.")

if __name__ == "__main__":
    main()
