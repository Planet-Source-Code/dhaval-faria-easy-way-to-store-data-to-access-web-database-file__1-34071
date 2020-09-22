
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
<%
dim FullName

FullName = Request.querystring
FullName = split(FullName,"=")
for i = 0 to ubound(FullName)
	FullName(i) = replace(FullName(i), "%20"," ")
next

dim cn
dim ReS

set cn = server.createobject("ADODB.Connection")
cn.Provider="Microsoft.Jet.OLEDB.4.0"
cn.Open Server.MapPath("Test.mdb")

set ReS = server.CreateObject("ADODB.Recordset")
ReS.Open "SELECT * from Test", Cn,1,3


res.AddNew
res("FullName") = FullName(1)
res("Gender") = FullName(3)
res("Age") = FullName(5)
res.update
%>
