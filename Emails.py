import pandas
import random
import time
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


password = input("Type your password and press enter:")
sender_email = "promo.writing.service@gmail.com"  

df = pandas.read_excel('D:\workspace Python\Emails\march.xlsx')
emails = df['emails'].values
#sites = df['site'].values


message = MIMEMultipart("alternative")
message["From"] = "Customer Support"   
#message["Subject"] = "🎁Find your discount at " + sites[0]
message["Subject"] = "🎁Find your discount at Marvelous-Essays💕" 


html = """
<p style="text-align: center;"><strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">Dear&nbsp;Customer,</span></strong></p>
<p style="text-align: center;"><strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">You&rsquo;ve got yourself an amazing deal at Marvelous - Essays!</span></strong></p>
<p style="text-align: center;"><strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">Here is&nbsp;</span></strong><strong><span style="font-size: 18.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">20% discount</span></strong><strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">&nbsp;with code:</span></strong><span style="font-size: 9.5pt; font-family: 'Century Gothic','sans-serif'; color: black;"> &nbsp;&nbsp;</span><u><span style="font-size: 28.0pt; font-family: 'Century Gothic','sans-serif'; color: #365f91;">20</span></u><u><span style="font-size: 28.0pt; font-family: 'Century Gothic','sans-serif'; color: #00b050;">mar</span></u></p>
<p style="margin: 0cm; margin-bottom: .0001pt; text-align: center;"><strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">Click&nbsp;</span></strong><span style="font-size: 14.0pt;"><a href="https://marvelous-essays.com/buy-essay.html"><strong><span style="font-family: 'Century Gothic','sans-serif'; color: #1155cc;">HERE&nbsp;</span></strong></a></span><strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">to order</span></strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">.</span></p>
<p style="text-align: center;"><strong><span style="font-size: 14.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">before it expires</span></strong> <strong><u><span style="font-size: 16.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">on May 15</span></u></strong><strong><span style="font-size: 16.0pt; font-family: 'Century Gothic','sans-serif'; color: #222222;">!</span></strong></p>
<p style="margin: 0cm; margin-bottom: .0001pt;"><span style="font-size: 7.0pt; font-family: 'Century Gothic','sans-serif'; color: #a6a6a6;">You received this message because you registered for an account on Marvelous-Essays</span></p>
<p><span style="font-size: 7.0pt; font-family: 'Century Gothic','sans-serif'; color: #a6a6a6;">Click here to&nbsp;</span><a href="mailto:support@marvelous-essays.com"><span style="font-size: 7.0pt; font-family: 'Century Gothic','sans-serif'; color: #cc0000;">unsubscribe</span></a></p>

"""
part1 = MIMEText(html, "html")    
message.attach(part1)



for i in range (len(emails)):
    #message["To"] = emails[i]    
    #receiver_email = emails[i]      
    #message["To"] = receiver_email     
    wait_time = random.randint(145,160) 
    print(i, wait_time, emails[i]) 
    #message["Subject"] = "🎁Find your discount at " + sites[i]
    
   
    
    

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)        
        server.sendmail(sender_email, emails[i], message.as_string())
        
    time.sleep (wait_time) 

print('DONE')