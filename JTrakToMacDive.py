#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement, Comment
from xml.dom import minidom

def prettify(elem):
    """Return a pretty-printed XML string for the Element.
    """
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")


# In[60]:


sheets = pd.read_excel("data/jtrak_dives3.xls",skiprows=3,sheet_name=None)


# In[61]:


sheets


# In[62]:


def seconds_from_mm(row):
    return sum(x * int(t) for x, t in zip([60, 1], str(row['Time']).split(":")))


# In[65]:


from lxml import etree
import datetime

def prettify_doc(rough_string):
    """Return a pretty-printed XML string for the Element.
    """
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

s = """<?xml version="1.0" encoding="UTF-8"?><dives />""".encode('utf-8')
parser = etree.XMLParser(ns_clean=True, recover=True, encoding='utf-8')
tree = etree.fromstring(s, parser=parser)
root = etree.Element("dives")


units = etree.SubElement(tree,'units')
units.text = 'Metric'

schema = etree.SubElement(tree,'schema')
schema.text = '2.2.0'
i=0
print(sheets['Jun 21, 2020 10_40 AM'])


# In[66]:


# sheets.pop('Jul 12, 2020 7:57 PM')
for sheet_name,dive in sheets.items():
    dive.dropna(inplace=True)
    
    dive['Time_seconds'] = dive.apply(seconds_from_mm, axis=1)
    dive_elem = etree.SubElement(tree,'dive')
    
    date = etree.SubElement(dive_elem,'date')
    date.text = datetime.datetime.strptime(sheet_name, '%b %d, %Y %I_%M %p').strftime('%Y-%m-%d %H:%M')
    
    identifier = etree.SubElement(dive_elem,'identifier')
    identifier.text = sheet_name
    
    diveNumber = etree.SubElement(dive_elem,'diveNumber')
    diveNumber.text = str(i)
    i += 1
    diver = etree.SubElement(dive_elem,'diver')
    diver.text = 'Claudio Spiess'
    
    computer = etree.SubElement(dive_elem,'computer')
    computer.text = 'Scubapro Aladin 2G'
    
    serial = etree.SubElement(dive_elem,'serial')
    serial.text = '9190'
    
    maxDepth = etree.SubElement(dive_elem,'maxDepth')
    maxDepth.text = "{:.1f}".format(dive['Depth [m]'].max())
    # avg depth dive['Depth [m]'].mean()
    # max depth dive['Depth [m]'].max()
    # first temp dive['Temp. [°C]'].iloc[0]
    # max temp dive['Temp. [°C]'].max()
    # min temp dive['Temp. [°C]'].min()
    averageDepth = etree.SubElement(dive_elem,'averageDepth')
    averageDepth.text = "{:.1f}".format(dive['Depth [m]'].mean())
    
    decoModel = etree.SubElement(dive_elem,'decoModel')
    decoModel.text = 'ZHL-8 ADT'
    
    duration = etree.SubElement(dive_elem,'duration')
    duration.text = "{:d}".format(len(dive))#dive["Time_seconds"].max())
    
    gasModel = etree.SubElement(dive_elem,'gasModel')
    gasModel.text = 'Air'
    
    sampleInterval = etree.SubElement(dive_elem,'sampleInterval')
    sampleInterval.text = '1'
    
    tempHigh = etree.SubElement(dive_elem,'tempHigh')
    tempHigh.text = "{:.1f}".format(dive["Temp. [°C]"].max())
    
    tempLow = etree.SubElement(dive_elem,'tempLow')
    tempLow.text = "{:.1f}".format(dive["Temp. [°C]"].min())
    
    gear = etree.SubElement(dive_elem,'gear')
    item = etree.SubElement(gear,'item')
    item_type = etree.SubElement(item,'type')
    item_type.text = 'Computer'
    manufacturer = etree.SubElement(item,'manufacturer')
    manufacturer.text = 'Scubapro'
    name = etree.SubElement(item,'name')
    name.text = 'Aladin 2G'
    serial = etree.SubElement(item,'serial')
    serial.text = '9190'
    
    samples = etree.SubElement(dive_elem,'samples')
    
    for index, row in dive.iterrows():
    #     print(row)
        sample = etree.SubElement(samples,'sample')
    
        time = etree.SubElement(sample,'time')
        time.text = "{:.1f}".format(index)
    
        depth = etree.SubElement(sample,'depth')
        depth.text = "{:.1f}".format(row["Depth [m]"])
    
        temperature = etree.SubElement(sample,'temperature')
        temperature.text = "{:.1f}".format(row["Temp. [°C]"])

#print(prettify_doc(etree.tostring(tree, encoding="UTF-8",
#                     xml_declaration=True,
#                     pretty_print=True,
#                     doctype='<!DOCTYPE dives SYSTEM "http://www.mac-dive.com/macdive_logbook.dtd">')))


# In[67]:


etree.ElementTree(tree).write('output.xml', pretty_print=True, xml_declaration=True,   encoding="utf-8",  doctype='<!DOCTYPE dives SYSTEM "http://www.mac-dive.com/macdive_logbook.dtd">')


# In[ ]:





# In[ ]:




