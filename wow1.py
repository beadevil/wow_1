from io import BytesIO
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import numpy as np
from streamlit_option_menu import option_menu
import plotly.express as px
from matplotlib import pyplot as plt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image
import requests


response = requests.get('https://static.mycareersfuture.gov.sg/images/company/logos/eb6e0f752982dff2188b6cbc89eef734/mindsprint.png')
image = Image.open(BytesIO(response.content))


st.sidebar.image(image, caption='MINDSPRINT')

st.sidebar.success("# `This Web App built in Python and Streamlit by  MAYUKH BHAUMIK 👨🏼‍💻`")

col1, col2, col3= st.columns(3)

with col1:
    st.title("MINDSPRINT")
with col2:
    st.write()

with col3:
    st.image(image, caption='MINDSPRINT')

st.write("")
st.write("")
st.write("")
st.write("")
st.write("")
st.write("")







excel_file = 'POC DATA.XLSX'
df = pd.read_excel(excel_file, sheet_name='MB DATA FINAL')

col = df.columns



with st.sidebar:
        selected = option_menu("Menu", ["Home", "MAIL"], 
        icons=['house', ], default_index=0)


if selected == "Home" :

    mat = list(df["Material Description"].unique())

    mat.insert(0, "SELECT")

    mat_des=st.sidebar.selectbox('SELECT Material :',mat)

    if mat_des != "SELECT":
        df = df[df["Material Description"].isin([mat_des])]


    df["Invoice Number"] = df["Invoice Number"].astype(str)

    plot_sel=st.sidebar.selectbox('SELECT PLOT :',["SELECT","SCATTER PLOT","BAR PLOT","PIE PLOT"])

    if plot_sel == "SELECT" :
        col_1, col_2= st.columns(2)

        with col_1:
            st.write("## Unique Products ")
            c=1
            for i in df["Material Description"].unique() :
                st.write(c,". ",i)
                c=c+1

    
        with col_2:
            for i in df["Material Description"].unique():
                st.write("## #. " ,i)
                dff = df[df["Material Description"].isin([i])]
                st.write("Min value of " , i , " : " , min(dff["Invoice Value- Net Price in USD"]))
                st.write("Max value of " , i , " : " , max(dff["Invoice Value- Net Price in USD"]))
                st.write("Avj value of " , i , " : " , sum(dff["Invoice Value- Net Price in USD"]) / len(dff))
                st.write("% value of " , i , "'s min and max differace : " , min(dff["Invoice Value- Net Price in USD"])*100/max(dff["Invoice Value- Net Price in USD"]))


    if plot_sel == "SCATTER PLOT":

        col = list(df.columns)

        col.insert(0, "SELECT")

        x= st.sidebar.selectbox('SELECT X Column for Bar Diagram :',col)


        y= st.sidebar.selectbox('SELECT Y Column for Bar Diagram :',col)

        warning_limit = st.sidebar.number_input('Enter an integer:', min_value=0, max_value=100, value=50, step=1)

        if x != "SELECT" and y != "SELECT" :

            x = np.array(df[x])
            y = np.array(df[y])

            fig, ax = plt.subplots(figsize=(12, 9))
            scatter = ax.scatter(x, y, c=y, cmap='coolwarm')

            cbar = plt.colorbar(scatter)
            cbar.set_label('Values')

            warning_limit = 0.5
            ax.axhline(y=warning_limit, color='orange', linestyle='--', label='Warning Limit')

            ax.set_xlabel('X')
            ax.set_ylabel('Y')
            ax.set_title('Scatter Plot')

            ax.tick_params(axis='x', labelrotation=90)

            st.pyplot(fig)


        else:
            for i in df["Material Description"].unique():
                st.write("## #. " ,i)
                dff = df[df["Material Description"].isin([i])]
                st.write("Min value of " , i , " : " , min(dff["Invoice Value- Net Price in USD"]))
                st.write("Max value of " , i , " : " , max(dff["Invoice Value- Net Price in USD"]))
                st.write("Avj value of " , i , " : " , sum(dff["Invoice Value- Net Price in USD"]) / len(dff))
                st.write("% value of " , i , "'s min and max differace : " , min(dff["Invoice Value- Net Price in USD"])*100/max(dff["Invoice Value- Net Price in USD"]))



    if plot_sel == "BAR PLOT":

        col = list(df.columns)

        col.insert(0, "SELECT")

        col_x= st.sidebar.selectbox('SELECT X Column for Bar Diagram :',col)


        col_y= st.sidebar.selectbox('SELECT Y Column for Bar Diagram :',col)

        warning_limit = st.sidebar.number_input('Enter an integer:', min_value=0, max_value=100, value=50, step=1)

        if col_x != "SELECT" and col_y != "SELECT" :

            categories = np.array(df[col_x])
            values = np.array(df[col_y])

            fig, ax = plt.subplots(figsize=(15, 15))

            ax.bar(categories, values)

            ax.plot(categories, values, marker='o', color='red', linewidth=2)

            ax.axhline(y=warning_limit, color='orange', linestyle='--', label='Warning Limit')
        
            ax.set_xlabel('Categories')
            ax.set_ylabel('Values')
            ax.set_title('Bar Plot with Line')

            ax.tick_params(axis='x', labelrotation=90)

            st.pyplot(fig)

        else:
            for i in df["Material Description"].unique():
                st.write("## #. " ,i)
                dff = df[df["Material Description"].isin([i])]
                st.write("Min value of " , i , " : " , min(dff["Invoice Value- Net Price in USD"]))
                st.write("Max value of " , i , " : " , max(dff["Invoice Value- Net Price in USD"]))
                st.write("Avj value of " , i , " : " , sum(dff["Invoice Value- Net Price in USD"]) / len(dff))
                st.write("% value of " , i , "'s min and max differace : " , min(dff["Invoice Value- Net Price in USD"])*100/max(dff["Invoice Value- Net Price in USD"]))



    if plot_sel == "PIE PLOT":
        col_p = list(df.columns)
        col_p.insert(0, "SELECT")
        col_pi = st.sidebar.selectbox('SELECT Column for Pie Diagram :',col_p)

        if col_pi != "SELECT":
            fig = plt.figure(figsize =(15, 15))
            unique, frequency = np.unique(np.array(df[col_pi]),
                              return_counts = True)

            plt.pie(frequency, labels=unique, autopct='%1.1f%%')

            plt.pie(frequency, labels = unique)

            plt.legend(title='Legend', loc='upper right' , labels=[f'{l}: {s}% ' for l, s in zip(unique, frequency)])

            st.pyplot(fig)

        else:
            for i in df["Material Description"].unique():
                st.write("## #. " ,i)
                dff = df[df["Material Description"].isin([i])]
                st.write("Min value of " , i , " : " , min(dff["Invoice Value- Net Price in USD"]))
                st.write("Max value of " , i , " : " , max(dff["Invoice Value- Net Price in USD"]))
                st.write("Avj value of " , i , " : " , sum(dff["Invoice Value- Net Price in USD"]) / len(dff))
                st.write("% value of " , i , "'s min and max differace : " , min(dff["Invoice Value- Net Price in USD"])*100/max(dff["Invoice Value- Net Price in USD"]))






if selected == "MAIL" :

    min_val = []
    max_val = []
    avj_val = []
    val_per = []

    df_avj = pd.DataFrame(columns = df.columns)

    per = st.sidebar.number_input('Enter limiter % :', min_value=0, max_value=100, value=60, step=5)

    for i in df["Material"].unique():
        dff = df[df["Material"].isin([i])]
        min_val.append(min(dff["Invoice Value- Net Price in USD"]))
        max_val.append(max(dff["Invoice Value- Net Price in USD"]))
        avj_val.append(sum(dff["Invoice Value- Net Price in USD"]) / len(dff))
        val_per.append(min(dff["Invoice Value- Net Price in USD"])*100/max(dff["Invoice Value- Net Price in USD"]))
        per_avj = (sum(dff["Invoice Value- Net Price in USD"]) / len(dff)) * per / 100
        df_avj_row = dff[dff['Invoice Value- Net Price in USD'] < per_avj]
        df_avj = pd.concat([df_avj, df_avj_row])
    
    st.write(df_avj)

    st.write("")
    st.write("")
    st.write("")
    st.success("DATA is ready to Downlode or Sending MAIL.")
    st.write("")


    def download_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.save()
        output.seek(0)
        return output

    button = st.download_button(label='Download Excel', data=download_excel(df_avj), file_name='data.xlsx')

    def gmail() :

        df_avj.to_excel('data.xlsx', index=False)

        try:
            s = smtplib.SMTP('smtp.gmail.com', 587)

            s.starttls()

            s.login("bhaumikmayukh@gmail.com", "qipywyfskjmofnky")

            sender_email = "bhaumikmayukh@gmail.com"
            recipient_email = "mayukh.bhaumik1999@gmail.com"
            cc_email = "bhaumikmayukh@gmail.com"
            subject = "Hello from Python!"
            body = "This is a test email sent from Python."
            attachment_path = "data.xlsx"

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Cc'] = cc_email
            msg['Subject'] = subject

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(attachment_path, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {attachment_path}")
            msg.attach(part)

            s.sendmail(sender_email, [recipient_email, cc_email], msg.as_string())

            s.quit()
            st.success("Email sent successfully!")
        except Exception as e:
            st.warning("Error: Unable to send email.")
            st.warning(e)
            st.warning("Please Send again")

    st.write("")
    st.write("")
    st.write("")
    

    if st.button(" SEND MAIL "):
        gmail()