import streamlit as st
import some_api_library

# Remove hardcoded credentials
# username = "login123"
# password = "password1233"

# Fetch credentials from Streamlit secrets
username = st.secrets["applogin"]["username"]
password = st.secrets["applogin"]["password"]

# Use credentials in the API login
api_client = some_api_library.Client(username=username, password=password)

# Rest of the code to interact with the API
response = api_client.do_something()
st.write(response)