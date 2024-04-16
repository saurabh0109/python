import smtplib
import pandas as pd

data = pd.read_csv("Excel.csv")
print(data.columns)
# print(data)

# creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)
# start TLS for security
s.starttls()
# Authentication
s.login("oxxxxxxxxxx@gmail.com", "xxxx xxxx xxxx xxxx")

i = 0
while i < 4:
    #print(i)
    acc_name = data['Account Name'][i]
    dmc_name = "" if pd.isnull(data['DMC NAME'][i]) else data['DMC NAME'][i]
    coex_name = "" if pd.isnull(data['COEX NAME'][i]) else data['COEX NAME'][i]
    acc_number = data['Account Number'][i]
    dmc_mail_id = data['DMC Email'][i]
    coex_mail_id = "" if pd.isnull(data['COEX Email'][i]) else data['COEX Email'][i]
    booking_center = data['Booking Centre'][i]
    assigned_to = data['Assigned To'][i]

    pending_items = data['Pending Items'][i].split(",")
    j = 0
    pending_items_str = ""
    while j < len(pending_items):
        k = j + 1
        pending_items_str += "\n" + str(k) + ". " + pending_items[j]
        j += 1

    # message to be sent
    coex_name_str = "" if not coex_name else " & " + coex_name
    addressed_to = dmc_name + coex_name_str
    subject_text = "Doc archival -" + str(booking_center) + " " + str(acc_number)
    body_text = "Dear " + addressed_to + ", \n\nRequesting your action on below items:\n" + pending_items_str + "\n\nRegards" + "\n" + assigned_to + "\nClient Formalities"
    to_list = [dmc_mail_id, coex_mail_id]
    message = 'Subject: {}\n\n{}'.format(subject_text, body_text)

    # sending the mail
    print(message)
    s.sendmail("oxxxxxxxxx@gmail.com", to_list, message)
    i += 1
# terminating the session
s.quit()