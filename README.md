# About JTrakToMacDive
This script converts JTrak XLS (Excel) sheets into MacDive's XML format.

# Motivation
I did my OW and AOWD with the dive shop's rental equipment, inc. computer. Unfortunately they only had the ancient Uwatec/Scubapro Aladin 2G, and the IrDA drivers only work on older Windows versions (I borrowed a friend's computer). I couldn't get SmartTRAK to work, but JTrak did. And unfortunately JTrak2XML didn't work. So I'm sharing the converter I wrote.

# Getting started
1. Export your JTrak logs to Excel XLS
2. Prepare your logs. Delete beginning/end rows that you don't want as part of your graph/dive length, and *rename* your sheets to include the date, the time currently must be split with an *_* instead of a *:*. If you used Numbers, you must export the sheets again as an Excel 2003 file (XLS).
3. Install the script:
```
git clone git@github.com:claudiosv/JTrakToMacDive.git
cd JTrakToMacDive
pip install pandas lxml
python xls2macdive.py path_to.xls
```
4. You should have an output.xml file in the same folder that you can open in MacDive.
