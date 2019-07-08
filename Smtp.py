import smtplib


class Smtp:
  def __init__(self, sender, passwd, smtp, port=587):
    self.email_sender = sender
    self.email_passwd = passwd
    self.smtp = smtp
    self.port = port

  def login(self):
    self.server = smtplib.SMTP(self.smtp, self.port)
    self.server.ehlo()
    self.server.starttls()
    self.server.login(self.email_sender, self.email_passwd)

  def send(self, to, subject, content):
    # try:
    body = '\r\n'.join(['To: %s' % to,
                        'From: %s' % self.email_sender,
                        'Subject: %s' % subject,
                        '', content])

    self.server.sendmail(self.email_sender, [to], body.encode("utf8"))

    # print('email sent')
    # except:
    #   print('error sending mail')

  def quit(self):
    self.server.quit()
