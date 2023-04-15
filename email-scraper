import imaplib
import email
from html2text import html2text
import pandas as pd
import re
import tkinter as tk

def email_reader(concat_string):
    # Connect to the server
    imap_server = imaplib.IMAP4_SSL('imap.gmail.com')

    # Login to the account
    # first you will have to configure Gmail's application password. details: https://red-box.readthedocs.io/en/latest/tutorials/config.html#gmail
    imap_server.login('abc@example.com', '{password}')

    # Select the mailbox you want to read from
    imap_server.select('INBOX')
    
    status, messages = imap_server.search(None, concat_string)

    extracted_area = []
    extracted_city = []
    extracted_state = []
    extracted_requirement = []
    extracted_email = []
    extracted_phone = []
    extracted_time = []

    # Loop through the messages
    for message in messages[0].split():
        # Fetch the message
        status, msg = imap_server.fetch(message, '(RFC822)')

        # Parse the message using the email library
        msg = email.message_from_bytes(msg[0][1])

        # Print the subject and sender of the message
        print(f'Subject: {msg["Subject"]}')
        print(f'From: {msg["From"]}')

        # Print the body of the message
        if msg.is_multipart():
            for part in msg.get_payload():

            
                if part.get_content_type() == "text/html":

                    body = html2text(part.get_payload(decode=True).decode(part.get_content_charset()))
               
                    user_area = re.findall(r'User Area: (.+)', body)
                    user_city = re.findall(r'User City: (.+)', body)
                    user_state = re.findall(r'User State: (.+)', body)
                    user_requirement = re.findall(r'User Requirement: (.+)', body)
                    user_time = re.findall(r'Search Date & Time: (.+)', body)
                    user_email = re.findall(r'User Email: (.+)',body)
                    user_phone = re.findall(r'User Phone: (.+)', body)
                
                    if not user_area and not user_city and not user_state and not user_requirement and (not user_email or not user_phone):
                        pass
                    else: 
                        #Extract User Area
                        if user_area: 
                            extracted_area.append(user_area)
                            print(user_area)
                        else: 
                            extracted_area.append("")
                            print("Text not Found!")
                
                        #Extract User City
                        if user_city:                    
                            extracted_city.append(user_city)
                            print(user_city)
                        else: 
                            extracted_city.append("")
                            print("Text not Found!")
                
                        #Extract User State
                        if user_state:                    
                            extracted_state.append(user_state)
                            print(user_state)
                        else: 
                            extracted_state.append("")
                            print("Text not Found!")
                
                        #Extract User Requirement
                        if user_requirement:                    
                            extracted_requirement.append(user_requirement)
                            print(user_requirement)
                        else: 
                            extracted_requirement.append("")
                            print("Text not Found!")
                
                        #Extract User Time
                        if user_time:                    
                            extracted_time.append(user_time)
                            print(user_time)
                        else: 
                            extracted_time.append("")
                            print("Text not Found!")
                    
                        #Extract User Email
                        if user_email:                    
                            extracted_email.append(user_email)
                            print(user_email)
                        else: 
                            extracted_email.append("")
                            print("Text not Found!")
                
                        #Extract User Phone Number
                        if user_phone:                    
                            extracted_phone.append(user_phone)
                            print(user_phone)
                        else: 
                            extracted_phone.append("")
                            print("Text not Found!")
                
        else:
        
            if msg.get_content_type() == "text/html":
                body = html2text(msg.get_payload(decode=True).decode(msg.get_content_charset()))

    data = {'Area': extracted_area,'City': extracted_city,'State':extracted_state,'Email':extracted_email,'Phone':extracted_phone,'Requirement': extracted_requirement}
    df = pd.DataFrame(data,dtype= str)
    df.to_excel('extracted_parts_2.xlsx', index=False)

    # Close the mailbox
    imap_server.close()

    # Logout from the account
    imap_server.logout()

window = tk.Tk()
window.geometry("500x300")
label = tk.Label(window, text="You want to search using subject or sender email: ")
subject_status = tk.Label(window, text="SEEN or UNSEEN Subject? ")
subject = tk.Label(window, text= "Enter the subject name: ")
sender_email = tk.Label(window,text= "Enter the Sender's Email address: ")
error_label = tk.Label(window,text= "Enter valid Criteria!! ",fg="red")

text_area = tk.Entry(window)
ss_text = tk.Entry(window)
subject_text = tk.Entry(window)
sender_email_text = tk.Entry(window)

def email_submit_button():
    
    email = sender_email_text.get()
    
    concatenated_string = f'FROM {email}'
    email_reader(concatenated_string)
    
    success = tk.Label(window, text= "Successfully added your data to excel file",fg="green")
    success.grid(row=6,column=1)
    
    sys.exit()
    

def subject_submit_button():
    
    ss = ss_text.get()
    sub = subject_text.get()
    
    concatenated_string = f'{ss.upper()} SUBJECT "{sub}"'
    print(concatenated_string)  
    email_reader(concatenated_string)
    
    success = tk.Label(window, text= "Successfully added your data to excel file",fg="green")
    success.grid(row=8,column=1)
    
    sys.exit()
     
def on_button_click():
    # Get the text from the Entry widget
    text = text_area.get()
    
    if text.upper() == "SUBJECT": 
        
        subject_status.grid(row=2,column=0)
        ss_text.grid(row=2,column=1)
        
        subject.grid(row=4,column=0)
        subject_text.grid(row=4,column=1)
        
        subject_button.grid(row=6,column=1)
        
    elif text.upper() == "SENDER" or text.upper() == "EMAIL" or text.upper() == "SENDER EMAIL":
        
        sender_email.grid(row=2,column=0)
        sender_email_text.grid(row=2,column=1)
        
        email_button.grid(row=4,column=1)
        
    else: 
        error_label.grid(row=2,column=1)
    
button = tk.Button(window, text="Get Text", command=on_button_click)
subject_button = tk.Button(window, text= "Submit Button", command= subject_submit_button)
email_button = tk.Button(window, text="Submit Button", command= email_submit_button)



label.grid(row=0,column=0)
text_area.grid(row=0,column=1)
button.grid(row=0,column=2)

window.title("Email Scraper")
window.mainloop()
