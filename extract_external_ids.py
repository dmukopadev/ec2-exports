import pandas as pd
import mysql.connector
import os
from datetime import datetime

def get_script_directory():
    """Get the directory where the current script is located"""
    return os.path.dirname(os.path.abspath(__file__))

def create_db_connection():
    """Create a connection to the MySQL database"""
    try:
        conn = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',
            database='primenet_pg'
        )
        print("✅ Successfully connected to the database")
        return conn
    except mysql.connector.Error as e:
        print(f"❌ Error connecting to database: {e}")
        return None

def read_excel_file(file_path):
    """Read the Excel file and return a DataFrame with proper date formatting"""
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"❌ Excel file not found: {file_path}")
            print(f"📁 Current directory: {os.getcwd()}")
            print(f"📁 Files in directory: {[f for f in os.listdir('.') if f.endswith('.xlsx')]}")
            return None
        
        df = pd.read_excel(file_path, sheet_name='Transactions')
        print(f"✅ Successfully read Excel file: {file_path}")
        print(f"📊 Found {len(df)} transactions in Excel file")
        
        # Convert the 'Created At' column to datetime format
        df['Created At'] = pd.to_datetime(df['Created At'], format='%d/%m/%Y, %I:%M:%S %p', errors='coerce')
        
        # Check for any date conversion errors
        if df['Created At'].isna().any():
            na_count = df['Created At'].isna().sum()
            print(f"⚠️  Warning: {na_count} dates could not be parsed properly")
        
        # Display basic info about the data
        print(f"📱 Unique phone numbers: {df['Phone Number'].nunique()}")
        print(f"💰 Amount range: {df['Amount'].min()} - {df['Amount'].max()}")
        print(f"📅 Date range: {df['Created At'].min()} - {df['Created At'].max()}")
        
        return df
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return None
    
def get_transactions_from_database(conn, excel_data):
    """Query database for transactions based on Excel data with exact date matching"""
    try:
        # Convert Excel data to formats suitable for SQL query
        account_numbers = excel_data['Account Number'].astype(str).tolist()
        amounts = excel_data['Amount'].tolist()
        
        # Convert dates to MySQL datetime format (YYYY-MM-DD HH:MM:SS)
        excel_dates = excel_data['Created At'].dt.strftime('%Y-%m-%d %H:%M:%S').tolist()
        
        print(f"🔍 Querying database for {len(account_numbers)} transactions...")
        print(f"📅 Sample dates from Excel: {excel_dates[:3]}")  # Show first 3 dates
        
        # Build the query to match transactions exactly by phone, amount, AND date
        query = """
        SELECT
            t.id,
            t.external_id,
            c.business_name AS Client,
            s.name AS Service,
            t.phone_number,
            t.account_number,
            t.narration,
            t.amount,
            t.created_at,
            t.updated_at,
            sc.status_name AS Status
        FROM transactions t
        JOIN clients c ON t.client_id = c.id
        JOIN status_codes sc ON t.status_code = sc.status_code
        JOIN services s ON t.service_id = s.id
        WHERE (t.account_number, DATE(t.created_at)) IN (
            {}
        )
        AND t.status_code = 300
        """.format(','.join(['(%s, %s, DATE(%s))'] * len(account_numbers)))
        
        # Create parameter list: [phone1, amount1, date1, phone2, amount2, date2, ...]
        params = []
        for i in range(len(account_numbers)):
            params.extend([account_numbers[i], excel_dates[i]])
        
        print(f"📝 Executing query with {len(params)} parameters...")
        
        # Execute query
        df_db = pd.read_sql(query, conn, params=params)
        print(f"✅ Found {len(df_db)} matching transactions in database")
        
        if not df_db.empty:
            print("\n📋 Sample of matched transactions:")
            print(df_db[['external_id', 'phone_number', 'amount', 'created_at']].head())
        else:
            print("❌ No transactions matched the criteria")
        
        return df_db
        
    except Exception as e:
        print(f"❌ Error querying database: {e}")
        return None

def generate_external_ids_excel(transactions_df, output_file):
    """Generate Excel file with only external_ids"""
    try:
        if transactions_df is None or transactions_df.empty:
            print("❌ No transactions found to export")
            return False
            
        # Check if external_id column exists
        if 'external_id' not in transactions_df.columns:
            print("❌ 'external_id' column not found in the results")
            print("Available columns:", transactions_df.columns.tolist())
            return False
        
        # Create DataFrame with only external_ids and remove duplicates
        external_ids_df = pd.DataFrame({
            'external_id': transactions_df['external_id'].drop_duplicates()
        })
        
        # Save to Excel
        external_ids_df.to_excel(output_file, index=False)
        print(f"✅ Successfully generated Excel file: {output_file}")
        print(f"📄 Exported {len(external_ids_df)} unique external_ids")
        
        # Display sample of exported external_ids
        print("\n📝 Sample of exported external_ids:")
        print(external_ids_df.head(10))
        
        return True
        
    except Exception as e:
        print(f"❌ Error generating Excel file: {e}")
        return False

def main():
    # Get the directory where the script is located
    script_dir = get_script_directory()
    
    # Configuration - files in the same directory as the script
    excel_file_path = os.path.join(script_dir, 'Primenet.xlsx')
    output_excel_path = os.path.join(script_dir, 'external_ids.xlsx')
    
    print("🚀 Starting transaction matching process...")
    print("=" * 50)
    print(f"📁 Script directory: {script_dir}")
    print(f"📊 Input file: {excel_file_path}")
    print(f"📤 Output file: {output_excel_path}")
    print("=" * 50)
    
    # Step 1: Read Excel file
    print("\n📖 Step 1: Reading Excel file...")
    excel_data = read_excel_file(excel_file_path)
    
    if excel_data is None:
        print("❌ Failed to read Excel file. Exiting.")
        return
    
    # Display sample of parsed dates
    print("\n📅 Sample of parsed data from Excel:")
    sample_data = excel_data[['Phone Number', 'Amount', 'Created At']].head()
    print(sample_data)
    
    # Step 2: Connect to database
    print("\n🔗 Step 2: Connecting to database...")
    conn = create_db_connection()
    
    if conn is None:
        print("❌ Failed to connect to database. Exiting.")
        return
    
    # Step 3: Query database for transactions
    print("\n🔍 Step 3: Querying database for matching transactions...")
    transactions = get_transactions_from_database(conn, excel_data)
    
    # Close database connection
    conn.close()
    print("✅ Database connection closed")
    
    if transactions is None or transactions.empty:
        print("❌ No matching transactions found in database.")
        print("💡 Possible reasons:")
        print("   - Dates/times don't match exactly between Excel and database")
        print("   - Phone number formats might be different")
        print("   - Amounts might have decimal differences")
        print("   - Transactions might have different status codes")
        return
    
    # Step 4: Generate Excel file with external_ids
    print("\n💾 Step 4: Generating Excel file with external_ids...")
    success = generate_external_ids_excel(transactions, output_excel_path)
    
    if success:
        print("\n" + "=" * 50)
        print("🎉 PROCESS COMPLETED SUCCESSFULLY!")
        print(f"📥 Input Excel: {excel_file_path}")
        print(f"📤 Output Excel: {output_excel_path}")
        print(f"📊 Total external_ids exported: {len(pd.read_excel(output_excel_path))}")
        print("=" * 50)
    else:
        print("\n❌ PROCESS FAILED!")

if __name__ == "__main__":
    main()