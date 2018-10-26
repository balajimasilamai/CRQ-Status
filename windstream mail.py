def send_mail(filename,to_addr):
    import smtplib
    import mimetypes
    import socks
    from email.mime.multipart import MIMEMultipart
    from email import encoders
    from email.message import Message
    from email.mime.audio import MIMEAudio
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email.mime.text import MIMEText
    #socks.setdefaultproxy(socks.HTTP, 'proxy.windstream.com', 8080)
    #socks.wrapmodule(smtplib)

    emailto = to_addr#"Balaji.Masilamani@windstream.com"
    emailfrom = 'CRQStatus@windstream.com'
    fileToSend = filename

    msg = MIMEMultipart()
    msg["From"] = emailfrom
    msg["To"] = emailto
    msg["Subject"] = "CRQ Status Report for date "
    msg.preamble = "CRQ Status Report"

    ctype, encoding = mimetypes.guess_type(fileToSend)
    if ctype is None or encoding is not None:
     ctype = "application/octet-stream"

    maintype, subtype = ctype.split("/", 1)

    if maintype == "text":
     fp = open(fileToSend)
    # Note: we should handle calculating the charset
     attachment = MIMEText(fp.read(), _subtype=subtype)
     fp.close()
    elif maintype == "image":
     fp = open(fileToSend, "rb")
     attachment = MIMEImage(fp.read(), _subtype=subtype)
     fp.close()
    elif maintype == "audio":
     fp = open(fileToSend, "rb")
     attachment = MIMEAudio(fp.read(), _subtype=subtype)
     fp.close()
    else:
     #with open(fileToSend,  'r', encoding='latin-1') as fp:
     fp = open(fileToSend, "rb")
     attachment = MIMEBase(maintype, subtype)
     attachment.set_payload(fp.read())
     fp.close()
     encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
    msg.attach(attachment)

    server = smtplib.SMTP("mailhost.windstream.com")
    #server.starttls()
    server.connect("mailhost.windstream.com")
    #server.login(username,password)
    try:
        server.sendmail(emailfrom, emailto,  msg.as_string())
    except Exception as e:
        messagebox.showinfo('Error', str(e) + '\n' + 'Check To address')
    server.quit()
send_mail('CRQ Status.xls','n9996094@windstream.com')
