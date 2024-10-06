#type:ignore
import pandas as pd
import numpy as nm 
from imap_tools import MailBox,AND,OR
import os
import streamlit as st
import datetime
from io import BytesIO

tailwind_css = """
<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
"""
st.markdown(tailwind_css, unsafe_allow_html=True)

class LoginToEmailUseImapTool:
    def __init__(self):
        self.__strings = []

    def getForm(self):
        with st.form("input_form"):
            lastSixDate = datetime.timedelta(days = 6)
            todayDate = datetime.date.today()
            lastDate = todayDate-lastSixDate
            self.select_option = st.multiselect('Select Option', ['Today Date','Form-To Date','Subject','From Email-Id','Search String'],key="select_option")
            self.subject = st.text_input('Subject Line',key="subject")
            self.today_date = st.date_input('Today Date',key="today_date")
            self.from_date = st.date_input('From Date',key="from_date",value=lastDate)
            self.to_date = st.date_input('To Date',key="to_date")
            self.search_string = st.text_input('Search Text',key="search_string")
            self.formSubmit = st.form_submit_button('Click To Get All Emails')
            
    def getExcelData(self):
        filePath = 'email.xlsx'
        if os.path.exists(filePath):
            self.data = pd.read_excel('email.xlsx')
            for i in self.data['search']:
                self.__strings.append(str(i).strip())
            self.__strings = nm.unique(tuple(self.__strings))

    def disable():
        return True

    def doLogin(self):
        try:
            with MailBox(f'imap.{self.provider.lower()}.com').login(self.email, self.password, initial_folder='INBOX') as self.mailbox:
                self.mailbox.folder.status('INBOX')
                print(f"Login Success by {self.email}")
                st.success(f"Welcome...!\n {self.email}.")# .You are logged successfully.
                self.getEmailDetails()
            return True
        except Exception as e:
            print("Mail Connection Issue:",e)
         
    def getLoginForm(self):
        # with st.spinner('Wait for it...'):
        #     time.sleep(5)
        # st.success("Done!")
        if "disabled" not in st.session_state:
            st.session_state.disabled = False

        with st.form("login_form"):
            self.provider = st.text_input('Provider(Ex. gmail,hotmail etc)',key="provider",value="gmail")
            self.email = st.text_input('Email-ID',key="email_id",value="shridip.chandole@gmail.com")
            self.password = st.text_input('Password',key="password",type="password", value="dime hpul ilux jnrb")
            self.loginSubmit = st.form_submit_button('Login',type="primary") #,disabled=st.session_state.disabled

        if(self.loginSubmit):
            with st.spinner("Logging in... Please wait..."):
                self.doLogin()
    
    def downloadExcelFile(self):
        dataFormat = {
            "search":["Amazon","SQL","Happy"]
        }
        st.write("'sample.xlsx' has been created,click to Download.")
        excelFileDataFrame = pd.DataFrame(dataFormat)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            excelFileDataFrame.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)
        processed_data = output.getvalue()
        st.download_button(label="Download Excel",data=processed_data,  file_name="sample.xlsx",mime="application/vnd.ms-excel")
    
    # upload excel file to search the data
    def uploadExcelFile(self):
         #d download sample.xlsx file here
        self.downloadExcelFile()
        self.uploadedFile = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"],accept_multiple_files=False,key="uploadedFile")
        self.email_date = st.date_input('Date',key="date",value=datetime.date.today())

        if self.uploadedFile:
            df = pd.read_excel(self.uploadedFile)
            st.write("File uploaded successfully!")
            print("File uploaded successfully!")
            self.__strings =[]
            # st.subheader("Excel Data")
            # st.dataframe(df)
            # st.subheader("File Details")
            # st.write(f"Number of rows: {df.shape[0]}")
            # st.write(f"Number of columns: {df.shape[1]}")
            for i in df.get('search'):
                self.__strings.append(str(i).strip())
            self.__strings = nm.unique(tuple(self.__strings))

    # get the emails of logged user
    def getEmailDetails(self):
        totalMailCount = 0
        try:
            #  below code is used to download the file
            # st.dataframe(self.data)
            # st.download_button(
            #         "Press to Download",
            #         self.data.to_excel(index=False).encode('utf-8'),
            #         "file.excel",
            #         "text/excel",
            #         key='download-csv'
            #         )
            # todayDate = self.today_date
            lastSixDate = datetime.timedelta(days = 6)
            
            todayDate = datetime.date.today()
            if(self.email_date):
               todayDate = self.email_date
            lastDate = todayDate-lastSixDate
            with st.expander(f"Here is your searchable keywords",expanded=False):
                    try:
                        if(len(self.__strings)):
                            st.write(f'''
                                <div class="shadow-sm pb-2"> {self.__strings}.</div>
                            ''',unsafe_allow_html=True)
                        else:
                            st.write(f'''
                                <div class="shadow-sm pb-2 text-yellow-600"> Search keyword not found</div>
                            ''',unsafe_allow_html=True)
                    except Exception as e:
                        print("Eror:",e)
            # lastDate = self.to_date
            # subject = self.subject
            # search_string = self.search_string
            # responses = mailbox.idle.wait(timeout=1)
            filePath = "Email Check.txt"
            if os.path.exists(filePath):
                os.remove(filePath)

            for searchString in self.__strings:
                for msg in self.mailbox.fetch(
                                        AND(text=f"%{searchString}%",
                                            date=todayDate
                                            # date_lt=todayDate,
                                            # date_gte=lastDate
                                            #,new=True
                                            #,subject="SQL"
                                            #,from_='from@ya.ru'
                                            # ,new=True
                                            )
                                        ):
                    st.markdown(f'''<div class="overflow-y-auto h-64 relative flex flex-col my-2 shadow-sm border border-slate-200 rounded-lg p-2">
                                        <p><b class=''>Subject:</b><span>{msg.subject}</span></p>
                                        <p><b>From:</b>{msg.from_}</p>
                                        <p><b>To</b>:{msg.to if(len(msg.to)!=0) else ''}</p>
                                        <p><b>CC</b>:{msg.cc if(len(msg.cc)!=0) else ''}</p>
                                        <p><b>BCC</b>:{msg.bcc if(len(msg.bcc)!=0) else ''}</p>
                                        <p><b>Date</b>:{msg.date_str}</p>
                                        <p><b>Body</b>:{msg.text}</p>
                                        <hr>
                                        <p class="text-center">*************** End Of the Email *******************</p>
                                    </div>
                                ''', unsafe_allow_html=True)
                    totalMailCount+=1

                    with open(filePath,'a') as file:
                        file.write(f"Subject:{msg.subject}\nDate:{msg.date_str}\nFrom:{msg.from_}\nTo:{msg.to}\nText:{msg.text}\n****************Yashodip******************\n")
        except Exception as e:
                print("Get Email details issue:",e)
        finally:
            print("Thanks all mail fetched...")
            with st.expander(f"Thanks all mail fetched..",expanded=False):
                    st.write(f'''
                        <div class="shadow-sm pb-2"> Total Emails are {totalMailCount}.</div>
                    ''',unsafe_allow_html=True)

def main():
    try:
        obj = LoginToEmailUseImapTool()
        # obj.getForm()
        obj.uploadExcelFile()
        obj.getLoginForm() 
        
    except Exception as e:
        print("Error:",e)
        return

if(__name__=="__main__"):
    main()
