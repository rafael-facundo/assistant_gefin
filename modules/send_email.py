from RPA.Email.ImapSmtp import ImapSmtp

gmail_account = "rfcosta@sfiec.org.br"
gmail_app_password = "jiitxvjdnglplqxx"
mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
mail.authorize(account=gmail_account, password=gmail_app_password)
mail.send_message(sender=gmail_account,
                  recipients=gmail_account,
                  subject="Email Test",
                  body="email test")