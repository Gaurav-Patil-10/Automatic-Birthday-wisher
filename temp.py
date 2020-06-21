import pandas as pd
import datetime
import smtplib

# ENTER THE AUTHENTICATION DETAILS
GMAIL_ID = '#############'
GMAIL_PASSWORD = '############'


def sendEmail(to, sub, message): # funtion for sending mails using the smtp library
    print(f"Email to {to} send with subject: {sub} and message {message} ")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PASSWORD)
    s.sendmail(GMAIL_ID , to , f"Subject: {sub}\n\n{message}")
    s.quit()



if __name__ == "__main__":
    
     # reading of the excel file which is containing the information of emails and dialogues
    df = pd.read_excel(
        r"C:\Users\User\Desktop\PROGRAMMING\PYTHON PROJECTS\Birthday Wisher\Birthday.xlsx")

    today = datetime.datetime.now().strftime("%d-%m")  # keeping the current time in date and month
    yearNow = datetime.datetime.now().strftime("%Y")
    write_ind = []
    for index, item in df.iterrows():  # itering the rows of the dataframe
        # print(index , item['Birthday'])
        bday = item["Birthday"].strftime("%d-%m")
        # print(bday)
        if today == bday and yearNow not in str(item['Year']):
            sendEmail(item["Emails"], "Happy Birthday", item["dialogue"])
            write_ind.append(index)
    # print(write_ind)

    for i in write_ind:    # writing the files for not repeating the Birthday wishes
        yr = df.loc[i, 'Year']
        # print(yr)
        df.loc[i, 'Year'] = str(yr) + ',' + str(yearNow)
        # print(df.loc[i, 'Year'])

    df.to_excel(r"C:\Users\User\Desktop\PROGRAMMING\PYTHON PROJECTS\Birthday Wisher\Birthday.xlsx", index=False)