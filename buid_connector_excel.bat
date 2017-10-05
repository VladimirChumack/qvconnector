REM Deleting Folders
RD /S /Q "dist"
RD /S /Q "build"

REM Deleting connenctor folder
RMDIR /s /q "C:\Program Files (x86)\Common Files\QlikTech\Custom Data\connector_excel"

REM Deleting Execution log
DEL /s /q connector.log

REM Building Connector
pyinstaller --version-file=file_version_info.txt --noupx --noconsole --icon=icon.ico connector_excel.py

REM Copying connector files across
xcopy dist\connector_excel  "C:\Program Files (x86)\Common Files\QlikTech\Custom Data\connector_excel\" /s /i
