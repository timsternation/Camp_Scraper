'''
Created on Feb 29, 2016

@author: timbo
'''

import openpyxl
from bs4 import BeautifulSoup #Needed to parse the html
from urllib.request import urlopen #needed to read the individual URLs
import time

wb = openpyxl.load_workbook('Camps.xlsx') #Starting file
sheet = wb.active

num = 4295 #Number pulled from the website

state_abb = [
             'AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL',
             'IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT',
             'NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI',
             'SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY','AS','DC',
             'FM','GU','MH','MP','PW','PR','VI'] # I created this list myself

for dig in range(1,num): #these next few lines contain the code for pulling the text down for each camp
    url = "http://find.acacamps.org/camp_profile.php?camp_id={}".format(dig)
    raw = urlopen(url)
    soup = BeautifulSoup(raw, 'lxml')
    contact = soup.findAll("div",class_="address_box")
    try: #There are a few dead links, so I used exception handling to skip these
        raw_text = contact[0].text
        cleaned_text = " ".join(raw_text.split())
        cleaner_text = cleaned_text.split()
        contact_text = cleaner_text[cleaner_text.index('Contact')+1:cleaner_text.index('Director:')]
        location_text = cleaner_text[cleaner_text.index('Location'):cleaner_text.index('Contact')]
        acc = soup.findAll('a', href="javascript:;", onclick="openStickyToolTip('accredited_member_help', 'below_origin', this)")

        #this gets whether the camp is accredited
        if len(acc) > 0:
            accm = "yes"
        else:
            accm = "no"

        name = []
        email = []

        if len(contact_text)<1:
    
            contact_email = "no email"
            contact_name = "no name"
    
        else:
            for element in contact_text:
                if '@' in element:
                    email.append(element)
        
                elif 'ext' in element == True or element[0] == 'x' or element[0].isdigit()==True or element[0]=="(" or 'www' in element or 'Camp' in element or '.com' in element or '.net' in element or '.org' in element:
                    continue
                else:
                    name.append(element)   
            
        print(name)
        print(email) 
   
        if len(name) == 0:
            contact_name = "no name"

        else:
            contact_name = " ".join(name)

        if len(email) == 0:
            contact_email = "no email"
    
        else:
            contact_email = " ".join(email)

        camp_name_raw = soup.findAll('h1')
        camp_name = camp_name_raw[0].text

        for entry in location_text:
            if len(entry)==2 and entry.isupper() == True and entry in state_abb:
                camp_state = entry
        

        print(camp_name)
        print(contact_name)
        print(contact_email)
        print(camp_state)

  
        sheet['A'+str(dig+1)] = camp_name
        sheet['B'+str(dig+1)] = contact_name
        sheet['C'+str(dig+1)] = contact_email
        sheet['D'+str(dig+1)] = camp_state
        sheet['E'+str(dig+1)] = accm  
        time.sleep(2) 
    except: #Skip a dead link- will appear as a blank row in the final document
        continue
        time.sleep(2)
        
wb.save('Camps2.xlsx')