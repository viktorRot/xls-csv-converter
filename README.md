# Excel to CSV converter
Convert an excel file to csv using apache poi library

## Build jar
```
mvn clean install
```

## Start application 
```
java -jar target\xls-csv-converter-0.0.1-SNAPSHOT-jar-with-dependencies.jar Skiprows=0 Seperator=; Sheetpos=0 sourcePath=C:\Users\ROTAEV\Desktop\temp\ archivePath=C:\Users\ROTAEV\Desktop\temp\Archive\
```

## The following arguments are used
```
###Optional:
Skiprows        Rows to skip, default 0
Seperator       Seperator for the columns, default ';'
Sheetpos        Sheet in the excel, which should be converted, default 0 (First sheet)

###Mandatory:
sourcePath      Path to the excel
archivePath     Path for archive the excel file after converting
```
