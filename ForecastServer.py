# -*- coding: utf-8 -*-
"""
Created on Tue Oct 11 17:44:10 2016

@author: alex.messina
"""

import urllib
from docx import *
from docx.shared import Inches
import datetime as dt



def Save_Forecast(site_name,url_list,maindir):
    ##  Create Document
    document = Document()
    ## Today's date
    now = str(dt.datetime.now())[:-13]+'00'
    ## Add  Heading 
    document.add_heading('NWS Forecast for '+site_name+' - '+now,level=1)
    ##
    num_count = 1
    for url in url_list:
        print url
        ## Get and Save Picture
        pic_filename = maindir+'NWS '+site_name+' Forecast'+str(num_count)+'.jpg'
        urllib.urlretrieve(url, pic_filename )
        ## Insert Picture
        document.add_picture(pic_filename ,width=Inches(7)) ## add pic from filename defined above
        ## count Up
        num_count += 1
    ## Save document
    document.save(maindir+'NWS Forecast for '+site_name+' - '+now+'.docx')
    return

## San Diego
maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/SanDiego/'
url_list = ['http://forecast.weather.gov/meteograms/Plotter.php?lat=32.7153&lon=-117.1573&wfo=SGX&zcode=CAZ043&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=32.7153&lon=-117.1573&wfo=SGX&zcode=CAZ043&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=32.7153&lon=-117.1573&wfo=SGX&zcode=CAZ043&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']
Save_Forecast('San Diego',url_list,maindir)

## Laguna Beach
maindir = 'C:/Users/alex.messina/Documents/GitHub/ForecastServer/LagunaBeach/'
url_list=['http://forecast.weather.gov/meteograms/Plotter.php?lat=33.5422&lon=-117.7831&wfo=SGX&zcode=CAZ552&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=0&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.5422&lon=-117.7831&wfo=SGX&zcode=CAZ552&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=48&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6', 'http://forecast.weather.gov/meteograms/Plotter.php?lat=33.5422&lon=-117.7831&wfo=SGX&zcode=CAZ552&gset=18&gdiff=3&unit=0&tinfo=PY8&ahour=96&pcmd=00000010100000000000000000000000000000000000000000000000000&lg=en&indu=1!1!1!&dd=&bw=&hrspan=48&pqpfhr=6&psnwhr=6']
Save_Forecast('LagunaBeach',url_list,maindir)





    