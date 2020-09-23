<div align="center">

## Display table from mdb as html


</div>

### Description

Often I use flat mdb files to post forms to on the webserver. This is a simple page to just grab a table from an mdb file on the server with a DSN-less connection and display it as an HTML table, so I can copy the table and paste it into Excel or Access, or just to get a quick look at the contents.
 
### More Info
 
Database file path, Table name

Shows use of ADO to get data and dynamically construct a table.

HTML table in browser of all columns / rows in an access db table.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jim Tinney](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jim-tinney.md)
**Level**          |Intermediate
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jim-tinney-display-table-from-mdb-as-html__4-7517/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<HTML><HEAD><META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
<TITLE>Database Table Contents</TITLE></HEAD><BODY>
<%
''Get the data from an access databse on the web server
''DSN-Less connection to local access db
dim Myconn,MyRS,MyField,MyRecord,DSNstr,SQLstr
	set conn = server.CreateObject("adodb.connection")
	set MyRS = server.CreateObject("adodb.recordset")
	DSNstr="DRIVER={Microsoft Access Driver (*.mdb)}; "
      DSNstr=DSNstr & "DBQ=" & "e:\wwwroot\databases\thing2.mdb"
''Change this to the path to your db file ^^^^^^^^^^^^^^^^^^^^^
  Myconn.Open DSNstr
   SQLstr = "SELECT * FROM results"
''Change this to your table name ^^^ (here it's "results")
	Set MyRS = Myconn.execute(SQLstr)
''You can send in the dbpath and tablename as parameters
''by simply calling the page like so:
''http://myserver/mydir/gettable.asp?db="e:\wwwroot\databases\thing2.mdb"&table="results"
''Then replace the strings above with
''    DSNstr=DSNstr & "DBQ=" & Request.Querystring("db")
''and
''    SQLstr = "SELECT * FROM " & Request.Querystring("table")
''This is not a very pretty table, meant mostly
''to get the table so you can copy it and paste
''it into Excel or something. Add formatting, widths,
''fonts, etc., to make it pretty.
response.Write "<table border=1><tr>"
MyRS.MoveFirst
For Each MyField In MyRS.Fields
 response.Write "<td>"
 response.Write MyField.Name
 response.Write "</td>"
Next 'MyField
response.Write "</tr>"
While not MyRS.EOF
 response.Write "<tr>"
  For Each MyField In MyRS.Fields
   response.Write "<td>"
   response.Write MyField.Value
   response.Write "</td>"
  Next 'MyField
 response.Write "</tr>"
 MyRS.MoveNext
Wend
response.Write "</table>"
''Close up objects
	Set MyRS= Nothing
	conn.Close
	set conn = nothing
''Simple, Huh?
%>
</BODY></HTML>
```

