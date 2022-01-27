<html>
<body>
<%

Dim adoCon 			'Holds the Database Connection Object
Dim rsGuestbook		'Holds the recordset for the records in the database
Dim strSQL			'Holds the SQL query for the database

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("DB.mdb")

'Create an ADO recordset object
Set rsGuestbook = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT Table1.* FROM Table1;"

'Open the recordset with the SQL query 
rsGuestbook.Open strSQL, adoCon

'Loop through the recordset
Do While not rsGuestbook.EOF
	
	'Write the HTML to display the current record in the recordset
	Response.Write ("<br>")
	Response.Write ("<a href=""Delete_entry.asp?ID=" & rsGuestbook("ID_no") & """>")
	Response.Write (rsGuestbook("Name")) 
	Response.Write ("</a>")
	Response.Write ("<br>")
	Response.Write (rsGuestbook("Email"))
	Response.Write ("<br>")
	Response.Write (rsGuestbook("Phone"))
	Response.Write ("<br>")
	Response.Write (rsGuestbook("Subject"))
	

	'Move to the next record in the recordset
	rsGuestbook.MoveNext

Loop

'Reset server objects
rsGuestbook.Close
Set rsGuestbook = Nothing
Set adoCon = Nothing
%>
</body>
</html>