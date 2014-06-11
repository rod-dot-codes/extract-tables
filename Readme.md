##Simple Python Write MSSQL to File
###Simply put, 
It is a simple way to extract data from a MSSQL database into Excel Tables.

First, ensure you got the requirements. (pip install -r requirements.txt)

Secondly,you define a SQL table to extract with JSON in a 'Tables' folder like this,


```json
 {
     "id": "Table",
     "server": "SERVER",
     "database": "DB",
     "query": "SELECT * FROM dbo.XYZ",
     "fields": ["COLUMN1","COLUMN2"],
     "mappings": ["Id","Name"]
 }
```

and thereafter place USER,PASS variable in secure.py (That's if like me, you haven't had time to setup domain authentication)

Then run 'python extract.py' and it should save it to Extracts\<id><date>.xls!

Hope you find it useful. Awesome to deal with pesky business analysts.
