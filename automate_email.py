from datetime import date, timedelta
import imaplib
import pandas as pd
import email
import yaml
import os
import os.path
import shutil

i = 0

today_date = date.today().strftime("%d %B %Y")

att_dir = today_date.replace(' ', '')

with open("credentials.yaml") as f:
    content = f.read()
credentials = yaml.load(content, Loader=yaml.FullLoader)
user, password, key, value, imap_url, mail_selector = credentials["user"], credentials["password"], credentials[
    "key"], credentials["value"], credentials["imap_url"], credentials["mail_selector"]


my_mail = imaplib.IMAP4_SSL(imap_url)

my_mail.login(user, password)

my_mail.select(mail_selector)

_, data = my_mail.search(None, key, value)
mail_id_list = data[0].split()

msgs = []

for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')
    msgs.append(data)

try:
    os.mkdir(att_dir)
except FileExistsError:
    shutil.rmtree(att_dir)
    os.mkdir(att_dir)

while i < 2:
    for msg in msgs[::-1]:
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg = email.message_from_bytes((response_part[1]))
                rcv_date = (date.today() - timedelta(i)).strftime("%d %B %Y")
                if rcv_date in my_msg["date"]:
                    print("_____________________________")
                    print("subj:", my_msg["subject"])
                    print("from:", my_msg["from"])
                    print("date:", my_msg["date"])
                    print("body:")
                    for part in my_msg.walk():
                        if part.get_content_type() == 'text/plain':
                            print(part.get_payload())
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        fileName = part.get_filename()
                        filePath = os.path.join(att_dir, fileName)
                        f_name = f"file{i}.xlsx"
                        xl_file = os.path.join(att_dir, f_name)

                        if bool(fileName):
                            with open(filePath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                                print(
                                    f"I'm writing the file {fileName} to disk {i} times")
                                f.close()
                        try:
                            os.rename(filePath, xl_file)
                        except FileExistsError:
                            os.remove(xl_file)
                            os.rename(filePath, xl_file)

    i += 1

xfile1 = pd.read_excel(os.path.join(att_dir, 'file0.xlsx'))

xfile2 = pd.read_excel(os.path.join(att_dir, 'file1.xlsx'))

xfile_join = xfile1.merge(right=xfile2, left_on=xfile1.columns.to_list(),
                          right_on=xfile2.columns.to_list(), how='outer')

xfile1.rename(columns=lambda x: x + '_file1', inplace=True)

xfile2.rename(columns=lambda x: x + '_file2', inplace=True)

xfile_join = xfile1.merge(right=xfile2, left_on=xfile1.columns.to_list(),
                          right_on=xfile2.columns.to_list(), how='outer')

xfile1_not_xfile2 = xfile_join.loc[xfile_join[xfile2.columns.to_list(
)].isnull().all(axis=1), xfile1.columns.to_list()]

xfile2_not_xfile1 = xfile_join.loc[xfile_join[xfile1.columns.to_list(
)].isnull().all(axis=1), xfile2.columns.to_list()]

xfile1_not_xfile2.to_excel(os.path.join(att_dir, 'new_items.xlsx'))

print(xfile1_not_xfile2)
