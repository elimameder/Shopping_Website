<%

Dim adoCon 			
Dim rsAddComments	
Dim strSQL			


Set adoCon = Server.CreateObject("ADODB.Connection")
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("DB.mdb")
Set rsAddComments = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT Table1.Name, Table1.Email, Table1.Phone, Table1.Subject FROM Table1;"


rsAddComments.CursorType = 2


rsAddComments.LockType = 3


rsAddComments.Open strSQL, adoCon


rsAddComments.AddNew

rsAddComments.Fields("Name") = Request.Form("Name")
rsAddComments.Fields("Email") = Request.Form("Email")
rsAddComments.Fields("Phone") = Request.Form("Phone")
rsAddComments.Fields("Subject") = Request.Form("Subject")

rsAddComments.Update

rsAddComments.Close
Set rsAddComments = Nothing
Set adoCon = Nothing

Response.Redirect "book.asp"


%>
