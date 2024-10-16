import os
import streamlit as st
import json
import datetime
import threading

# Authentication data
username_password = {'admin': 'password123'}

# Authentication function
def authenticate(username, password):
    return username in username_password and username_password[username] == password

# Function to run the data extraction script
def run_extraction_script():
    try:
        os.system('python "C:\\Users\\Rolando Pancho\\Documents\\EXTRACTEXCELFILEFINAL.py"')
        last_extraction_date = datetime.datetime.now().date()
        with open('last_extraction.json', 'w') as f:
            json.dump({'last_extraction': last_extraction_date.strftime('%Y-%m-%d')}, f)
        st.success("Data extraction completed, check SharePoint.")
        update_last_extraction_info()
    except Exception as e:
        st.error(f"Error: {e}")

# Update the last extraction info
def update_last_extraction_info():
    try:
        with open('last_extraction.json', 'r') as f:
            data = json.load(f)
            last_extraction_date = datetime.datetime.strptime(data['last_extraction'], '%Y-%m-%d').date()
            days_since = (datetime.datetime.now().date() - last_extraction_date).days
            st.write(f"Last extraction date: {last_extraction_date} ({days_since} days ago)")
    except (FileNotFoundError, json.JSONDecodeError):
        st.write("No previous extraction found.")

# Main Streamlit application
def main():
    st.title("Greenprofi Data Extraction Tool")

    # Authentication form
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.sidebar.header("Login")
        username = st.sidebar.text_input("Username", key="username")
        password = st.sidebar.text_input("Password", type="password", key="password")
        if st.sidebar.button("Login", key="login"):
            if authenticate(username, password):
                st.session_state.authenticated = True
                st.sidebar.success("Login successful!")
                st.experimental_rerun()
            else:
                st.sidebar.error("Invalid username or password")
        st.warning("Please log in using the sidebar.")
        return

    # Display last extraction info
    update_last_extraction_info()

    # Extract Data button
    if st.button("Extract Data"):
        threading.Thread(target=run_extraction_script).start()

    # Logout button
    if st.sidebar.button("Logout"):
        st.session_state.authenticated = False
        st.experimental_rerun()

if __name__ == "__main__":
    main()