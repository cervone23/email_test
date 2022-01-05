import streamlit as st 

import win32com.client as client

if st.button("Get EMAIL"):

    number = st.slider("Pick a number",0, 100)

    outlook = client.Dispatch("Outlook.Application")

    message = outlook.CreateItem(0)

    message.Display()

if __name__ == '__main__':
    main()