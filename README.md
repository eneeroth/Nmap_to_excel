# How To Use
The script listens for xml files, you can input single, multiple or wildcard (*.xml).
```
python3 nmap_to_excel.py *.xml -o <outputname>.xlsx
```
or
```
python3 nmap_to_excel.py <file1>.xml <file2>.xml <fileX>.xml -o <outputname>.xlsx
```

# Limitations
At the moment the script is created for host discovery, portscans and/or version scans to simplify the report process. Further nmap outout may work but has not been tested.