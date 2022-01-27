<%
'Dimension variables
Dim adoCon			'Holds the Database Connection Object
Dim rsDeleteEntry	'Holds the recordset for the record to be deleted
Dim strSQL			'Holds the SQL query for the database
Dim lngRecordNo		'Holds the record number to be deleted

'Read in the record number to be deleted
lngRecordNo = CLng(Request.QueryString("ID"))

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("DB.mdb")

'Create an ADO recordset object
Set rsDeleteEntry = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT Table1.* FROM Table1 WHERE ID_no=" & lngRecordNo

'Set the lock type so that the record is locked by ADO when it is deleted
rsDeleteEntry.LockType = 3

'Open the recordset with the SQL query 
rsDeleteEntry.Open strSQL, adoCon

'Delete the record from the database
rsDeleteEntry.Delete

'Reset server objects
rsDeleteEntry.Close
Set rsDeleteEntry = Nothing
Set adoCon = Nothing

'Return to the delete select page incase another record needs deleting
Response.Redirect "Delete_select.asp"
%>