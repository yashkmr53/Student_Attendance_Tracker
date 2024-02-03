import pandas as pd 
import smtplib

mail_id='Enter your mail id here'
password= 'Enter your password here'

def sendEmail (to,msg,sub):
    s=smtplib.SMTP(host="smtp.gmail.com",port=587)
    s.starttls()
    s.login(user=mail_id,password=password)
    s.sendmail(from_addr=mail_id, to_addrs=to, msg=f"Subject : {sub}\n\n {msg}")
    s.quit()
    pass

if __name__=="__main__":
    df=pd.read_excel("studentattendence.xlsx")
    subCode=input("Enter the Subject Code : ").upper().strip()
    absentRoll=input("Enter the roll number of absent students : ").upper().split()
    for i in absentRoll:
        for j in range((df.shape)[0]):
            if i==df.iloc[j,2]:
                df.loc[j,subCode]=df.loc[j,subCode]+1
                if df.loc[j,subCode]==1:
                    to=df.loc[j,"Email"]
                    msg=f"Only 1 leave is now left in the subject code : {subCode}"
                    sub="Warning"
                    sendEmail(to,msg,sub)
                if df.loc[j,subCode]>=2:
                    to=df.loc[j,"Email"]
                    msg=f"You have exhausted your all leaves in the subject code : {subCode}"
                    sub="Warning"
                    sendEmail(to,msg,sub)
                df.to_excel("studentattendence.xlsx",index=False)
                