import streamlit as st
import os
import json
import datetime
import threading

def run_script():
    try:
        os.system('python "C:\\Users\\Rolando Pancho\\Documents\\EXTRACTEXCELFILEFINAL.py"')
        last_extraction_date = datetime.datetime.now().date()
        with open('last_extraction.json', 'w') as f:
            json.dump({'last_extraction': last_extraction_date.strftime('%Y-%m-%d')}, f)
        st.success("Data extraction completed, check SharePoint.")
    except Exception as e:
        st.error(f"Error: {e}")

def main():
    st.title("Greenprofi Data Extraction Tool")

    try:
        with open('last_extraction.json', 'r') as f:
            data = json.load(f)
            last_extraction_date = datetime.datetime.strptime(data['last_extraction'], '%Y-%m-%d').date()
            days_since = (datetime.datetime.now().date() - last_extraction_date).days
            st.write(f"Last extraction date: {last_extraction_date} ({days_since} days ago)")
    except (FileNotFoundError, json.JSONDecodeError):
        st.write("No previous extraction found.")

    if st.button("Extract Data"):
        threading.Thread(target=run_script).start()

if __name__ == "__main__":
    main()
