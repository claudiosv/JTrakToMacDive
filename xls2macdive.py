#!/usr/bin/env python
# coding: utf-8
import pandas as pd
from lxml import etree
import datetime


def seconds_from_mm(row):
    return sum(x * int(t) for x, t in zip([60, 1], str(row['Time']).split(":")))

xls_location = "data/jtrak_dives3.xls"

sheets = pd.read_excel(xls_location, skiprows=3, sheet_name=None)

s = """<?xml version="1.0" encoding="UTF-8"?><dives />""".encode('utf-8')
parser = etree.XMLParser(ns_clean=True, recover=True, encoding='utf-8')
tree = etree.fromstring(s, parser=parser)
root = etree.Element("dives")

units = etree.SubElement(tree, 'units')
units.text = 'Metric'

schema = etree.SubElement(tree, 'schema')
schema.text = '2.2.0'
i = 1

for sheet_name, dive in sheets.items():
    dive.dropna(inplace=True)

    dive['Time_seconds'] = dive.apply(seconds_from_mm, axis=1)
    dive_elem = etree.SubElement(tree, 'dive')

    date = etree.SubElement(dive_elem, 'date')
    date.text = datetime.datetime.strptime(
        sheet_name, '%b %d, %Y %I_%M %p').strftime('%Y-%m-%d %H:%M')

    identifier = etree.SubElement(dive_elem, 'identifier')
    identifier.text = sheet_name

    diveNumber = etree.SubElement(dive_elem, 'diveNumber')
    diveNumber.text = str(i)
    i += 1

    maxDepth = etree.SubElement(dive_elem, 'maxDepth')
    maxDepth.text = "{:.1f}".format(dive['Depth [m]'].max())

    averageDepth = etree.SubElement(dive_elem, 'averageDepth')
    averageDepth.text = "{:.1f}".format(dive['Depth [m]'].mean())

    duration = etree.SubElement(dive_elem, 'duration')
    duration.text = "{:d}".format(len(dive))

    sampleInterval = etree.SubElement(dive_elem, 'sampleInterval')
    sampleInterval.text = '1'

    tempHigh = etree.SubElement(dive_elem, 'tempHigh')
    tempHigh.text = "{:.1f}".format(dive["Temp. [°C]"].max())

    tempLow = etree.SubElement(dive_elem, 'tempLow')
    tempLow.text = "{:.1f}".format(dive["Temp. [°C]"].min())

    samples = etree.SubElement(dive_elem, 'samples')

    for index, row in dive.iterrows():
        sample = etree.SubElement(samples, 'sample')

        time = etree.SubElement(sample, 'time')
        time.text = "{:.1f}".format(index)

        depth = etree.SubElement(sample, 'depth')
        depth.text = "{:.1f}".format(row["Depth [m]"])

        temperature = etree.SubElement(sample, 'temperature')
        temperature.text = "{:.1f}".format(row["Temp. [°C]"])


etree.ElementTree(tree).write('output1.xml', pretty_print=True, xml_declaration=True,   encoding="utf-8",
                              doctype='<!DOCTYPE dives SYSTEM "http://www.mac-dive.com/macdive_logbook.dtd">')
