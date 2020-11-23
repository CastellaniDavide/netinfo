# netinfo
[![GitHub license](https://img.shields.io/badge/licence-GNU-green?style=flat)](https://github.com/CastellaniDavide/cpp-netinfo/blob/master/LICENSE) ![Author](https://img.shields.io/badge/author-Castellani%20Davide-green?style=flat) ![Version](https://img.shields.io/badge/version-v05.01-blue?style=flat) ![Language Python](https://img.shields.io/badge/language-Python-yellowgreen?style=flat) ![sys.platform supported](https://img.shields.io/badge/OS%20platform%20supported-Windows-blue?style=flat) [![On GitHub](https://img.shields.io/badge/on%20GitHub-True-green?style=flat&logo=github)](https://github.com/CastellaniDavide/netinfo)

## Tags
 #wmi, #automation, #python3, #network, #ping, #csv, #database, #datetime

## Description
Get net infos in all computers in "computers.csv" file & in "computers.db" database

## Required
 - python3
 - pip3 librerias (in project core: pip3 install -r requirements/requirements.txt)

## Goals
 - [x] Save in csv file the PC's infos
 - [x] Save the infos in a database
 - [x] Print only if in debug mode
 - [x] Make standard csv
 - [x] Make database optional
 - [x] Check if PC is avariable (using ping)
 - [x] Use append method
 - [x] Add date info
 - [x] Add unchecked_PC.csv file and unchecked_PC table for the PCs that i can't reach
 - [x] Add folder variable
 - [x] Add Vendom info
 - [x] Normalize MACAddress

## Directories structure
 - .github
   - ISSUE_TEMPLATE
     - bug_report.md
     - feature-request.md
 - bin
   - **netinfo.py**
 - docs or doc
   - LICENSE
   - README.md
   - _config.yml
 - flussi
   - computers.csv
   - netinfo.csv <- On first run
   - netinfo.db (optional) <- On first run
   - netinfo_history.csv <- On first run
   - unchecked_PC.csv
 - log
   - log.log
 - requirements
   - requirements.txt
   
### Execution examples
 1. update "computers.csv" file with our PC
 2. (Clear "log.log" file)
 3. **python3 netinfo.py**

# Changelog
 - [Version_05.01-2020-11-23](#Version_0501-2020-11-23)
 - [Version_04.01-2020-11-09](#Version_0401-2020-11-09)
 - [Version_03.03_2020-10-15](#Version_0303_2020-10-15)
 - [Version_03.02_2020-10-15](#Version_0302_2020-10-15)
 - [Version_03.01_2020-10-15](#Version_0301_2020-10-15)
 - [Version_02.01_2020-09-28](#Version_0201_2020-09-28)
 - [Version_01.01_2020-09-21](#Version_0101_2020-09-21)

## Version_05.01-2020-11-23
 - Normalize MACAddress

## Version_04.01-2020-11-09
 - Fixed a bug
 - Add Vendom info (eg. Productor name)

## Version_03.03_2020-10-15
 - Fixed a bug

## Version_03.02_2020-10-15
 - trace.log -> log.log

## Version_03.01_2020-10-15
 - Fixed some bugs
 - Add folder variable

## Version_02.01_2020-09-21
 - Fixed some bugs
 - Optimized csv
 - Print only if in debug mode
 - Made database optional
 - Check if PC is avariable (using ping)
 - Use append method
 - Add date info
 - Add unchecked_PC.csv file and unchecked_PC table for the PCs that i can't reach
 - Added a WebSite with the last version

## Version_01.01_2020-09-21
 - Initial version

---
Made by Castellani Davide 
If you have any problem please contact me:
 - [help@castellanidavide.it](mailto:help@castellanidavide.it)
 - [Issue](https://github.com/CastellaniDavide/netinfo/issues)
