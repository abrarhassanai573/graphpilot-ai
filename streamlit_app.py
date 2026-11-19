import streamlit as st

st.set_page_config(
    page_title="GraphPilot AI - Diagnostic Test",
    page_icon="🧭",
    layout="wide",
)

st.title("GraphPilot AI")
st.success("Streamlit interface is rendering correctly.")
st.write("This is a temporary diagnostic page.")
st.info("If you can read this page, the basic Streamlit frontend is working.")

with st.sidebar:
    st.header("Diagnostic")
    st.write("Sidebar is visible.")
