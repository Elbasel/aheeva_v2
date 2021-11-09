import os
from time import sleep

try:
    os.chdir(r'Z:\WFM Data\RTM Shared Folder\steramlit_dashboard')
    os.system('streamlit run streamlit_dashboard.py')
    input()
except Exception as ex:
    print(ex)
    input()
