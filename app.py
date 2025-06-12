{\rtf1\ansi\ansicpg1252\cocoartf2822
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 import streamlit as st\
import pandas as pd\
import openai\
\
st.set_page_config(page_title="Excel Chatbot", layout="centered")\
\
st.title("\uc0\u55357 \u56522  Ask Questions About Your Excel File")\
\
openai.api_key = st.secrets["OPENAI_API_KEY"]\
\
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])\
\
if uploaded_file:\
    df = pd.read_excel(uploaded_file)\
    st.write("Preview:")\
    st.dataframe(df.head())\
\
    question = st.text_input("Ask your question")\
\
    if question:\
        context = df.head(20).to_string(index=False)\
        prompt = f"""You are a data analyst. Given this data:\\n\\n\{context\}\\n\\nAnswer this: \{question\}"""\
\
        with st.spinner("Thinking..."):\
            response = openai.ChatCompletion.create(\
                model="gpt-3.5-turbo",\
                messages=[\{"role": "user", "content": prompt\}]\
            )\
            answer = response.choices[0].message.content.strip()\
            st.success("Answer:")\
            st.write(answer)\
}