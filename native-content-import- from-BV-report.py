import sys
import xlrd
import datetime
from lxml import objectify, etree as ET
import pandas as pd


def main():
    file = sys.argv[1]
    df = pd.read_excel(file)
    f = pd.DataFrame(df)
    n = open("migration.xml", "w")
    n.write('<?xml version="1.0" encoding="UTF-8"?>')
    n.close()
    root = ET.Element("Feed")
    root.set("xmlns", "http://www.bazaarvoice.com/xs/PRR/StandardClientFeed/14.7")
    root.set("name", sys.argv[2])
    root.set('extractDate', datetime.datetime.now().isoformat())
    product = ET.SubElement(root, "Product")
    product.set('id', '20202106')
    eid = ET.SubElement(product, "ExternalId")
    eid.text = "20202106"
    rws = ET.SubElement(product, "Reviews")
    for index, row in f.iterrows():
        rw = ET.SubElement(rws, "Review")
        rw.set('id', str(row['Review ID']))
        rw.set('removed', 'false')
        mds = ET.SubElement(rw, "ModerationStatus")
        mds.text= str(row['Moderation Status'])
        usr = ET.SubElement(rw, 'UserProfileReference')
        usr.set('id', str(row['User ID']))
        uid = ET.SubElement(usr, "ExternalId")
        uid.text = str(row['User ID'])
        dn = ET.SubElement(usr, "DisplayName")
        dn.text = str(row['User Nickname'])
        anon = ET.SubElement(usr, "Anonymous")
        hyp = ET.SubElement(usr, 'HyperlinkingEnabled')
        anon.text = 'false'
        hyp.text = 'false'
        title = ET.SubElement(rw, "Title")
        title.text = str(row['Review Title'])
        rwtxt = ET.SubElement(rw, "ReviewText")
        rwtxt.text = str(row['Review Text'])
        rating = ET.SubElement(rw, "Rating")
        rating.text = str(row['Overall Rating'])
        recom = ET.SubElement(rw, "Recommended")
        if row['Recommend to a Friend'] == "Yes":
            recom.text = 'true'
        else:
            recom.text = 'false'
        email = ET.SubElement(rw, "UserEmailAddress")
        email.text=str(row['User Email Address'])
        loc = ET.SubElement(rw, "ReviewerLocation")
        loc.text = str(row['User Location'])
        sub = ET.SubElement(rw, "SubmissionTime")
        sub.text = str(row['Submission Date'])[:10]+'T'+str(row['Submission Date'])[11:]+".0000"
        ftr = ET.SubElement(rw, 'Featured')
        ftr.text = 'false'
        locale = ET.SubElement(rw, 'DisplayLocale')
        locale.text = str(row['Locale'])

    xmldata = ET.tostring(root)
    n = open("migration.xml", 'ab')
    n.write(xmldata)
    n.close()



if __name__ == "__main__":
    try:
        main()
    except Exception as inst:
        print(type(inst))    # the exception instance
        print(inst)
