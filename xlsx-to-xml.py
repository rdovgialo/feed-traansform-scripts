import sys
import datetime
from lxml import objectify, etree as ET
import pandas as pd
import random


file = sys.argv[1]
newFile = file.replace(".xlsx", ".xml")
df = pd.read_excel(file)
f = pd.DataFrame(df)

root = ET.Element("Feed")
root.set("xmlns", "http://www.bazaarvoice.com/xs/PRR/StandardClientFeed/14.7")
root.set('name', sys.argv[2])
root.set('extractDate', datetime.datetime.now().isoformat())

users = {}

for index, row in f.iterrows():
    if row['ProductID']:
        product = ET.SubElement(root, "Product")
        product.set('id', str(row['ProductID'])[:-2])
        pid = ET.SubElement(product, "ExternalId")
        pid.text = str(row['ProductID'])[:-2]
        rws = ET.SubElement(product, "Reviews")
        rw = ET.SubElement(rws, "Review")
        rw.set('id', str(row['ReviewsID']))
        #rw.set('removed', 'false')
        usr = ET.SubElement(rw, "UserProfileReference")

        uid = ET.SubElement(usr, "ExternalId")
        if str(row['UserDisplayName']) in users:
            uid.text = users[str(row['UserDisplayName'])]
            usr.set('id', users[str(row['UserDisplayName'])])
        else:
            users[str(row['UserDisplayName'])] = str(random.randint(100, 5000))
            uid.text = users[str(row['UserDisplayName'])]
            usr.set('id', users[str(row['UserDisplayName'])])
        uname = ET.SubElement(usr, "DisplayName")
        uname.text = str(row['UserDisplayName'])
        anon = ET.SubElement(usr, "Anonymous")
        anon.text = "false"
        hpr = ET.SubElement(usr, "HyperlinkingEnabled")
        hpr.text = "false"
        Text = ET.SubElement(rw, "ReviewText")
        Text.text = str(row['ReviewText'])
        rating = ET.SubElement(rw, "Rating")
        rating.text = str(row['Rating'])
        stime = ET.SubElement(rw, "SubmissionTime")
        stime.text = f"{str(row['SubmissionTime'])[:10]}T{str(row['SubmissionTime'])[11:]}.000000"
        dloc = ET.SubElement(rw, "DisplayLocale")
        dloc.text = str(row['DisplayLocale'])


f = open(newFile, "w")
f.write('<?xml version="1.0" encoding="UTF-8"?>\r\n')
newxml = ET.tostring(root, pretty_print=True)
f.close()
f = open(newFile, 'ab')
f.write(newxml)
f.close()
