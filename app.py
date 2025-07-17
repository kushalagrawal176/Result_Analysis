import streamlit as st
from SE import *
from TE import *
from BE import *
from config import getConfig

st.title("Result Analysis")
st.header("1. Upload Files")

file1 = st.file_uploader("Curent year Excel File", type=["xlsx", "xls"], key = "1")
file2 = st.file_uploader("Previous year Excel File", type=["xlsx", "xls"], key = "2")

year = st.selectbox("select year", ["SE", "TE", "BE"])
semester = st.selectbox("select semester", ["I", "II"])

if file1 and file2:
    if st.button("Process File"):
        if(year == 'SE'):
            if(semester == "I"):
                sub = getConfig("SEM-III")
            else:
                sub = getConfig("SEM-IV")

            with st.spinner("Processing..."):
                result = SE_analysis(file1, file2, sub)

            st.success("File processed successfully!")
            st.download_button(
                label="Download Processed File",
                data=result,
                file_name="Result_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif(year == "TE"):
            if(semester == "I"):
                sub = getConfig("SEM-V")
            else:
                sub = getConfig("SEM-VI")

            with st.spinner("Processing..."):
                result = TE_analysis(file1, file2, sub)

            st.success("File processed successfully!")
            st.download_button(
                label="Download Processed File",
                data=result,
                file_name="Result_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            if(semester == "I"):
                sub = getConfig("SEM-VII")
            else:
                sub = getConfig("SEM-VIII")

            with st.spinner("Processing..."):
                result = BE_analysis(file1, file2, sub)

            st.success("File processed successfully!")
            st.download_button(
                label="Download Processed File",
                data=result,
                file_name="Result_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )