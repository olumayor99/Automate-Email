from datetime import date, timedelta
import imaplib
import pandas as pd
import email
import yaml
import os
import os.path

i = 0

today_date = date.today().strftime("%d %B %Y")

attachment_dir = today_date.replace(' ', '')
print(f'{attachment_dir}/file1.xlsx')

with open("credentials.yaml") as f:
    content = f.read()
my_credentials = yaml.load(content, Loader=yaml.FullLoader)
user, password, key, value, imap_url, mail_selector = my_credentials["user"], my_credentials["password"], my_credentials[
    "key"], my_credentials["value"], my_credentials["imap_url"], my_credentials["mail_selector"]


my_mail = imaplib.IMAP4_SSL(imap_url)

my_mail.login(user, password)

my_mail.select(mail_selector)

_, data = my_mail.search(None, key, value)
mail_id_list = data[0].split()

msgs = []

for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')
    msgs.append(data)

for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            while (True):

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
                        filePath = os.path.join(attachment_dir, fileName)
                        f_name = f"file{i+1}.xlsx"
                        xl_file = os.path.join(attachment_dir, f_name)

                        if bool(fileName):
                            try:
                                os.mkdir(attachment_dir)
                            except Exception:
                                pass
                            with open(filePath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                        try:
                            os.rename(filePath, xl_file)
                        except FileExistsError:
                            os.remove(xl_file)
                            os.rename(filePath, xl_file)
                if i == 2:
                    break

                i += 1

df1 = pd.read_excel(f'{attachment_dir}/file1.xlsx')

df2 = pd.read_excel(f'{attachment_dir}/file1.xlsx')

df_join = df1.merge(right=df2, left_on=df1.columns.to_list(),
                    right_on=df2.columns.to_list(), how='outer')

df1.rename(columns=lambda x: x + '_file1', inplace=True)

df2.rename(columns=lambda x: x + '_file2', inplace=True)

df_join = df1.merge(right=df2, left_on=df1.columns.to_list(),
                    right_on=df2.columns.to_list(), how='outer')

records_present_in_df1_not_in_df2 = df_join.loc[df_join[df2.columns.to_list(
)].isnull().all(axis=1), df1.columns.to_list()]

records_present_in_df2_not_in_df1 = df_join.loc[df_join[df1.columns.to_list(
)].isnull().all(axis=1), df2.columns.to_list()]

print(records_present_in_df1_not_in_df2)
