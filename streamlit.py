import pandas as pd
import streamlit as st
import pywhatkit as kit
import datetime
import time

try:
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    df = pd.read_excel('Jan 2024.xlsx', sheet_name='2\'MAR24')
except:
    st.write("Please upload the excel file")
else:
    def generate_alphabetical_names(n):
        alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        result = []
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result[:0] = alphabet[remainder]
        return ''.join(result)

    # Assign alphabetical names to columns
    df.columns = [generate_alphabetical_names(i + 1) for i in range(len(df.columns))]

# if 'packing' not in st.session_state:
#     st.session_state.expanded = False

# if st.button('Packing aid'):
#     st.session_state.expanded = not st.session_state.expanded
# , expanded=st.session_state.expanded
with st.expander("Packing aid"):
    c=0
    for index, row in df.iterrows():
        tot=0
        msg=''
        phone=''
        area=row['E']
        if not pd.isna(row['C']):
            phone='+91 '+str(int(row['C']))
        if(type(row['C'])==float and phone !=''):
            st.markdown(f"## {row['D']}  {area}")
            for column, value in row.items():
                vegetable=df.at[2,column]
                price=df.at[1,column]
                if not pd.isna(value) and not pd.isna(price) and not pd.isna(vegetable):
                    tot+=float(value)*float(price)
                    st.markdown(f"##### {vegetable} ----- {value} ")

            comments=row['CB']
            st.markdown(f"##### Total Amount: {tot}\n\n")
            st.markdown(f"##### comments--> {comments}\n\n")
if (st.button('Send bill')):
  with st.expander("Send bill"):
    # BILL SCRIPT
    import pywhatkit as kit
    import datetime
    import time
    c=0
    for index, row in df.iterrows():
        tot=0
        msg=''
        phone=''
    #     print(row['H'],row['E'])
        if not pd.isna(row['C']):
            phone='+91 '+str(int(row['C']))
            print(phone)
        if(type(row['C'])==float and phone !=''):
            msg+=f"ANNAM FARMS\t{datetime.date.today()}\nHi {row['D']},\nYour bill details are:\n"
            for column, value in row.items():
                vegetable=df.at[2,column]
                price=df.at[1,column]
                if not pd.isna(value) and not pd.isna(price) and not pd.isna(vegetable):
                    msg+=f"\n{vegetable}= {value} *{price}={round(float(value)*float(price))}"
                    tot+=float(value)*float(price)
            delivery=df.columns[df.iloc[2].eq('Delivery')][0]
    #         comments=df.columns[df.iloc[5].eq('Comments')][0]
            tot+=df.at[index, delivery]
    #         msg+=f"\nDelivery={df.at[index, delivery]}\n{df.at[index, comments]}\nTotal={tot}"
    #         if not pd.isnull(df.at[index, comments]):
    #             msg += f"\nDelivery={df.at[index, delivery]}\n{df.at[index, comments]}\nTotal={round(tot)}"
    #         else:
            msg += f"\nDelivery={df.at[index, delivery]}\nTotal={round(tot)}"

        print(msg)
        
        try:
            if msg:
                c += 1
                kit.sendwhatmsg_instantly(phone, msg)
                time.sleep(7)
        except Exception as e:
            print(f"Error: Message not sent to {phone}. Error details: {str(e)}")
        print("-----------------")
    print(c)  