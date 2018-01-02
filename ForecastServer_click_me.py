# -*- coding: utf-8 -*-
"""
Created on Tue Oct 11 17:44:10 2016

@author: alex.messina
"""



if __name__=='__main__':
    import urllib
    from docx import *
    from docx.shared import Inches
    import datetime as dt
    import time    
    
    from email.MIMEMultipart import MIMEMultipart
    from email.MIMEText import MIMEText
    from email.MIMEImage import MIMEImage
    import smtplib
    
    
    def Save_Forecast(site_name,url_list,maindir,email=False):
        print 'Looking up weather forecast for '+site_name+'....'
        print
        ##  Create Document
        document = Document()
        ## Create Email
        msg = MIMEMultipart()
        ## Today's date
        now = str(dt.datetime.now())[:-13]+'00'
        ## Add  Heading to Document
        header = 'NWS Forecast for '+site_name+' - '+now
        document.add_heading(header,level=1)
        ## Add Heading to email
        msg['SUBJECT']=header
        msg.attach(MIMEText(header))
        ##
        num_count = 1
        for url in url_list:
            #print url
            ## Get and Save Picture
            pic_filename = maindir+'NWS '+site_name+' Forecast'+str(num_count)+'.png'
            urllib.urlretrieve(url, pic_filename )
            ## Insert Picture
            document.add_picture(pic_filename ,width=Inches(7)) ## add pic from filename defined above
            ## Add Picture to email
            msg.attach(MIMEImage(file(pic_filename,'rb').read()))
            ## count Up
            num_count += 1
            
        ## Save document
        print 'Saving forecasts...'
        document.save(maindir+'NWS Forecast for '+site_name+' - '+now+'.docx')
        if email==True:
            ## Send Email
            # to send
            mailer = smtplib.SMTP_SSL('smtp.googlemail.com:465')
            username, password = 'sebal.py', 'Mactec101'
            mailer.login(username,password)
            mailer.sendmail('sebal.py@gmail.com','atm1984@gmail.com', msg.as_string())
            mailer.close()
        print
        print 'Forecasts saved uh successfurry . Have a nice day :)'
        print
        return
    
    
    ## San Diego
    maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/SanDiego/'
    maindir = 'SanDiego/'
    SanDiego_url_list = ['http://forecast.weather.gov/meteograms/Plotter.php?lat=32.7153&lon=-117.1573&wfo=SGX&zcode=CAZ043&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=32.7153&lon=-117.1573&wfo=SGX&zcode=CAZ043&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=32.7153&lon=-117.1573&wfo=SGX&zcode=CAZ043&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']
    
    ## Laguna Beach
    maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/LagunaBeach/'
    maindir = 'LagunaBeach/'
    LagunaBeach_url_list=['http://forecast.weather.gov/meteograms/Plotter.php?lat=33.5422&lon=-117.7831&wfo=SGX&zcode=CAZ552&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.5422&lon=-117.7831&wfo=SGX&zcode=CAZ552&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.5422&lon=-117.7831&wfo=SGX&zcode=CAZ552&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']
    
    ## March AFB/Riverside
    maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/Riverside/'
    maindir = 'Riverside/'
    Riverside_url_list=['http://forecast.weather.gov/meteograms/Plotter.php?lat=33.899&lon=-117.2522&wfo=SGX&zcode=CAZ048&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.899&lon=-117.2522&wfo=SGX&zcode=CAZ048&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.899&lon=-117.2522&wfo=SGX&zcode=CAZ048&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']

    ## Avalon, Catalina Island
    maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/Avalon/'
    maindir = 'Avalon/'
    Avalon_url_list=['http://forecast.weather.gov/meteograms/Plotter.php?lat=33.3428&lon=-118.3278&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.3428&lon=-118.3278&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.3428&lon=-118.3278&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']

    ## Dana Point
    maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/DanaPoint/'
    maindir = 'DanaPoint/'
    DanaPoint_url_list=['http://forecast.weather.gov/meteograms/Plotter.php?lat=33.4805&lon=-117.6977&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.4805&lon=-117.6977&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.4805&lon=-117.6977&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']

    ## John Wayne Airport
    maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/JohnWayne/'
    maindir = 'JohnWayne/'
    JohnWayne_url_list=['http://forecast.weather.gov/meteograms/Plotter.php?lat=33.688&lon=-117.9043&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.688&lon=-117.9043&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.688&lon=-117.9043&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']
    
    ## Foothill Ranch
    maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/FoothillRanch/'
    maindir = 'FoothillRanch/'
    FoothillRanch_url_list=['http://forecast.weather.gov/meteograms/Plotter.php?lat=33.688&lon=-117.9043&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.688&lon=-117.9043&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.688&lon=-117.9043&wfo=LOX&zcode=CAZ087&gset=15&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']



    ## For testing only:
    #Save_Forecast('Avalon',Avalon_url_list,maindir='C:/Users/alex.messina/Documents/GitHub/ForecastServer/Avalon/',email=True)

#    Save_Forecast('San Diego',SanDiego_url_list,maindir='SanDiego/',email=True)
#    Save_Forecast('LagunaBeach',LagunaBeach_url_list,maindir='LagunaBeach/',email=True)
#    Save_Forecast('Riverside',Riverside_url_list,maindir='Riverside/',email=True)
#    Save_Forecast('Avalon',Avalon_url_list,maindir='Avalon/',email=True)
#    Save_Forecast('DanaPoint',Avalon_url_list,maindir='DanaPoint/',email=True)
#    Save_Forecast('JohnWayne',JohnWayne_url_list,maindir='JohnWayne/',email=True)
    Save_Forecast('FoothillRanch',FoothillRanch_url_list,maindir='FoothillRanch/',email=True)
    print 'Waiting for next forecast download interval...'

    print str(dt.datetime.now())
 

    #print 'Press any key to exit....'
    #input()
    print 'Forecasts saved.'



    
