# MINI DOCUMENTATION / DESCRIPTION
#
# Headers taht are used in Product feed are hardcoded
# hard coded headers : ['Marsha', 'Location Name', 'Property description short', 'Brand', 'Description - AU', 'Description - CN', 'Description - DE', 'Description - ES', 'Description - FR', 'Description - JP', 'Description - UK', 'Name - AU', 'Name - CN', 'Name - DE', 'Name - ES', 'Name - FR', 'Name - JP', 'Name - PT', 'Name - UK', 'URL - AU', 'URL - CN', 'URL - DE', 'URL - ES', 'URL - FR', 'URL - JP', 'URL - PT', 'URL - UK', 'ProductURL' ]
# All other headers will be added as custom Attributes
# Script logs:
#   * Products that are missing a defout PDP URL
#   * Product that are missing a brand name
#   * Python exceptions
#   * Type of Python exceptions
import sys
import xlrd
import pandas as pd
import datetime
from lxml import etree as ET

def main():
    file = sys.argv[1]
    newFile = "productfeed.xml"

    df = pd.read_excel(file)
    f = pd.DataFrame(df)
    n = open(newFile, 'w')
    # to save space in memmory lxml will reate only one product at a time and append that to newFile
    allHeaders = list(f.columns.values)
    PFHeaders = ['Marsha', 'Location Name', 'Property description short', 'Brand', 'Description - AU', 'Description - CN', 'Description - DE', 'Description - ES', 'Description - FR', 'Description - JP', 'Description - UK', 'Name - AU', 'Name - CN', 'Name - DE', 'Name - ES', 'Name - FR', 'Name - JP', 'Name - PT', 'Name - UK', 'URL - AU', 'URL - CN', 'URL - DE', 'URL - ES', 'URL - FR', 'URL - JP', 'URL - PT', 'URL - UK', 'ProductURL' ]
    #initiating new file as xml
    n.write('<?xml version="1.0" encoding="UTF-8"?>')
    #adding top level elements Product and Feed
    n.write(f'<Feed extractDate="{datetime.datetime.now().isoformat()}" incremental="false" name="{sys.argv[2]}" xmlns="http://www.bazaarvoice.com/xs/PRR/ProductFeed/14.5"><Products>')
    n.close()
    for index, row in f.iterrows():
        root = ET.Element("Product")
        pid = ET.SubElement(root, "ExternalId")
        pid.text = row['Marsha']
        nameDefault = ET.SubElement(root, "Name")
        nameDefault.text = row['Location Name']
        names = ET.SubElement(root, "Names")
        # checking if has lokilized data:
        if isinstance(row['Name - AU'], str):
            en_AU_name = ET.SubElement(names, "Name")
            en_AU_name.set('locale', 'en_AU')
            en_AU_name.text = row['Name - AU']
        if isinstance(row['Name - CN'], str):
            ch_CN_name = ET.SubElement(names, "Name")
            ch_CN_name.set('locale', 'ch_CN')
            ch_CN_name.text = row['Name - CN']
        if isinstance(row['Name - DE'], str):
            de_DE_name = ET.SubElement(names, "Name")
            de_DE_name.set('locale', 'de_DE')
            de_DE_name.text = row['Name - DE']
        if isinstance(row['Name - ES'], str):
            es_ES_name = ET.SubElement(names, "Name")
            es_ES_name.set('locale', 'es_ES')
            es_ES_name.text = row['Name - ES']
        if isinstance(row['Name - FR'], str):
            fr_FR_name = ET.SubElement(names, "Name")
            fr_FR_name.set('locale', 'fr_FR')
            fr_FR_name.text = row['Name - FR']
        if isinstance(row['Name - JP'], str):
            ja_JP_name = ET.SubElement(names, "Name")
            ja_JP_name.set('locale', 'ja_JP')
            ja_JP_name.text = row['Name - JP']
        if isinstance(row['Name - PT'], str):
            pt_PT_name = ET.SubElement(names, "Name")
            pt_PT_name.set('locale', 'pt_PT')
            pt_PT_name.text = row['Name - PT']
        if isinstance(row['Name - UK'], str):
            en_GB_name = ET.SubElement(names, "Name")
            en_GB_name.set('locale', 'en_GB')
            en_GB_name.text = row['Name - UK']
        # Remove "Names" where no lokalized data is present
        if len(names) == 0:
            names.getparent().remove(names)
        if isinstance(row['Property description short'], str):
            defaultDescription = ET.SubElement(root, "Description")
            defaultDescription.text = row['Property description short']
        if isinstance(row['Brand'], str):
            brand = ET.SubElement(root, "Brand")
            brandName = ET.SubElement(brand, "Name")
            brandName.text = row['Brand']
        else:
            print(f"{row['Marsha']} is missing Brand\n")
        if isinstance(row['ProductURL'], str):
            productPageUrl = ET.SubElement(root, "ProductPageUrl")
            productPageUrl.text = row['ProductURL']
        else:
            print(f"{row['Marsha']} is missing Default product page URL\n")
        PDPs = ET.SubElement(root, "ProductPageUrls")
        if isinstance(row['URL - AU'], str):
            en_AU_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            en_AU_PDP.set('locale', 'en_AU')
            en_AU_PDP.text = row['URL - AU']
        if isinstance(row['URL - CN'], str):
            ch_CN_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            ch_CN_PDP.set('locale', 'ch_CN')
            ch_CN_PDP.text = row['URL - CN']
        if isinstance(row['URL - DE'], str):
            de_DE_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            de_DE_PDP.set('locale', 'de_DE')
            de_DE_PDP.text = row['URL - DE']
        if isinstance(row['URL - ES'], str):
            es_ES_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            es_ES_PDP.set('locale', 'es_ES')
            es_ES_PDP.text = row['URL - ES']
        if isinstance(row['URL - FR'], str):
            fr_FR_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            fr_FR_PDP.set('locale', 'fr_FR')
            fr_FR_PDP.text = row['URL - FR']
        if isinstance(row['URL - JP'], str):
            ja_JP_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            ja_JP_PDP.set('locale', 'ja_JP')
            ja_JP_PDP.text = row['URL - JP']
        if isinstance(row['URL - PT'], str):
            pt_PT_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            pt_PT_PDP.set('locale', 'pt_PT')
            pt_PT_PDP.text = row['URL - PT']
        if isinstance(row['URL - UK'], str):
            en_GB_PDP = ET.SubElement(PDPs, "ProductPageUrl")
            en_GB_PDP.set('locale', 'en_GB')
            en_GB_PDP.text = row['URL - UK']
        if len(PDPs) == 0:
            PDPs.getparent().remove(PDPs)

        #Adding custom attributes
        attributes = ET.SubElement(root, "Attributes")
        for i in allHeaders:
            if i not in PFHeaders:
                 if isinstance(row[i], str):
                     attr = ET.SubElement(attributes, "Attribute")
                     attr.set('id', i.replace(" ", "_"))
                     value = ET.SubElement(attr, "Value")
                     value.text = row[i]


        if len(attributes) == 0:
            attributes.getparent().remove(attributes)

        xmldata = ET.tostring(root)
        n = open(newFile, 'ab')
        n.write(xmldata)
        n.close()
    # append the closing tags for top level elements Products and Feed
    n = open(newFile, 'a')
    n.write('</Products></Feed>')
    n.close()

#for debuging running main funciton in  try / except, on exception log the exeoption, this will be seen in stdrtErr in AWS
if __name__ == "__main__":
    try:
        main()
    except Exception as inst:
        print(type(inst))    # the exception instance
        print(inst)
