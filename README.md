# Automate-Email

Make sure you complete the following steps before using this repo

1. Go to your Gmail settings for the account you want to get the excel files from.

2. Click on "Forwarding and POP/IMAP" and tick "Enable IMAP".

3. Go to the Google Account settings of that same account, click on "Security", turn on 2FA if it isn't on, than click on "App Passwords".

4. Click on "Select app", and in the dropdown menu, select mail. In the "Select device" dropdown, select "Other (Custon name)".

5. Input a custom name, then click on "Generate" which will generate a password. Save it somewhere.

6. Go to the credentials file and set the user as the Gmail address, and the password as the generated password from step 5. You can also use it as sender_address and senser_pass for the SMTP server.

7. Create an EC2 instance (you can modify the script for any platform), ssh into it and run the following commands:

   - sudo apt update && sudo apt upgrade -y

   - sudo apt install software-properties-common -y

   - sudo add-apt-repository ppa:deadsnakes/ppa (press ENTER to continue)

   - sudo apt install python3.10

   - sudo apt install python3-pip

   - pip install pandas openpyxl

8. After filling the credentials.yaml file, copy credentials.yaml, email_auto.py, email_auto.sh, and email_auto.service to the /home/ubuntu/ folder of the EC2 instance.

9. Move "email_auto.service" to /etc/systemd/system/ folder, then run "sudo chmod 664 /etc/systemd/system/email_auto.service"

10. Run "chmod 744 /home/ubuntu/email_auto.sh" or "chmod u+x email_auto.sh"

11. Run "sudo systemctl daemon-reload", then "sudo systemctl enable email_auto.service"

12. Reboot your instance afterwards to confirm the script is working.

13. Create a cron job in the EC2 instance by running "crontab -e" as you wish. I used "0 12 \* \* 0-6 /usr/bin/python3 /home/ubuntu/email_auto.py" which means "At 12:00 on every day-of-week from Sunday through Saturday, run email_auto.py". You can create custom cron jobs using https://crontab.guru/.

Feel free to ask me any additional questions.
