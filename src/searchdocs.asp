<%@ Language=VBScript %>
<!--#include file="../_private/adovbs.inc" -->
<%
Const RecordsPerPage = 10   ' Numarul de inregistrari de pe o pagina

printpageheader             ' Se afiseaza headerul paginii

if Request.QueryString.Count=0 then 
  Response.Write "<center><font size=+2 color=#000090><b>VMA HardDisk Search Engine</b></font></center>"
  printsearchform "",""
else
  qe=Request.QueryString("q")
  nqe=Request.QueryString("nq")
  printsearchform qe,nqe
  DoSearch qe,nqe
end if  

'
' Aceasta este subrutina principala ce face interogarea bazei
' lui Index Server si apoi afiseaza paginat rezultatele
'
sub DoSearch(q,nq)
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.ConnectionString = "provider=msidxs"
  objConn.Open
  Set objRS = Server.CreateObject("ADODB.RecordSet")
  objRS.CursorLocation = adUseClient

  on error resume next
  objRS.Open BuildSQLQuery(q,nq), objConn, adOpenKeyset,adLockReadOnly
  If Err.number <> 0 or (objRS.EOF and objRS.BOF) Then
    Err.Clear
    Response.Write "No records found!"
    Set objRS = Nothing
    objConn.Close
    Set objConn = Nothing
    Response.End 
    Exit Sub
  End If

  objRS.PageSize = RecordsPerPage
  pag = Request.QueryString("pag")
  If pag <> "" Then
    pag = CInt(pag)
    If pag < 1 Then pag = 1
  Else
    pag = 1
  End If
  objRS.AbsolutePage = pag
  
  RowCount = objRS.PageSize
  RowOnPage = RecordsPerPage * (pag-1) +1
  Response.Write "<table border=0 width=90% align=center cellpadding=10>"
  
  Response.Write "<tr><td colspan=2 bgcolor=#e0e0e0>"
  Response.Write "S-au gasit " & objRS.RecordCount & " documente ce contin: <b>"&q&"</b>"
  if nq<>"" then Response.Write " si nu contin <b>"&nq&"</b>"
  Response.Write "</td></tr>"
  
  Do While Not objRS.EOF And RowCount > 0 
    Response.Write "<tr>"
    Response.Write "<td valign=top align=left width=10>"
    Response.Write RowOnPage & "."
    RowOnPage = RowOnPage + 1
    with Response
     .Write "</td>"
     .Write "<td valign=top align=left>"
     .Write "<b>Title: " & Server.HTMLEncode(objRS.Fields("DocTitle").Value) & "</b><br>"
     .Write "<b>Filename:</b> " & objRS.Fields("Filename").Value & "  - " & objRS.Fields("size").Value & " bytes - " &objRS.Fields("write").Value & "<br>"
     .Write "<b>Description:</b> " & Server.HTMLEncode(objRS.Fields("characterization").Value) & "<br>"
     .Write "<b>URL: </b><a href='" & objRS.Fields("path").Value & "'>"
     .Write objRS.Fields("path").Value
     .Write "</a><br>"&vbcrlf&vbcrlf&vbcrlf
     .Write "</td>"
     .Write "</tr>"
    end with
    RowCount = RowCount - 1 
    objRS.MoveNext
  loop
  
  Response.Write "<tr><td colspan=2><table width=100% cellspacing=0 cellpadding=5 border=0 bgcolor=#e0e0e0 align=center><tr>"
  
  If pag > 1 Then
   Response.Write "<td align=left><a href='searchdocs.asp?q="&q&"&nq="&nq&"&pag="&pag-1&"'>Back</a></td>"
  end if 
  If RowCount = 0 Then 
   Response.Write "<td align=right><a href='searchdocs.asp?q="&q&"&nq="&nq&"&pag="&pag+1&"'>Next</a></td>"
  End If

  Response.Write "</tr></table></td></tr>"
  Response.Write "</table>"

 Set objRS = Nothing
 objConn.Close
 Set objConn = Nothing
end sub


'
' Construieste interogarea SQL pentru Index Server
'
function BuildSQLQuery(q,nq)
 SQL = "SELECT Rank, Filename, Size, DocTitle, Path, Write, Characterization FROM System..Scope() "&_
       "WHERE CONTAINS(" & "'" & q & "'" & ")"
 if nq<>"" then SQL = SQL & " AND NOT CONTAINS(" & "'" & nq & "'" & ")"
 SQL = SQL & " ORDER BY Rank DESC"
 BuildSQLQuery = SQL      
end function 


'
' Tipareste headerul paginii
'
sub printpageheader%>
<head>
<title>VMA HardDisk Search Engine</title>
<style>
body,td
 {
   font-family:verdana;
   font-size:10;
 }
a.cool
 {
   color:#000090;
   text-decoration:none;
 }
a.cool:hover
 {
   color:red;
   text-decoration:underline;
 }  
</style>
</head>
<%end sub

'
' Subrutina tipareste formul de search din prima pagina
' si din partea de sus a paginilor cu rezultate
'
sub printsearchform(q,nq)

if nq<>"" then
  searchformt1 = ""
  searchformt2 = "Simple<br>mode"
  searchformt3 = ""
else
  searchformt1 = " display:none;"
  searchformt2 = "Advanced<br>mode"
  searchformt3 = " style='display:none;'"
end if
%>

<script language=vbscript>
sub switchmodetext_onclick
 if tablerow2.style.display="none" then
   containstext.style.display=""
   switchmodetext.innerhtml="Simple<br>mode"
   tablerow2.style.display=""
 else
   containstext.style.display="none"
   switchmodetext.innerhtml="Advanced<br>mode"
   tablerow2.style.display="none"
 end if 
end sub

sub seachform_onsubmit
 if tablerow2.style.display="none" then
   seachform.nq.value = ""
 end if
end sub

sub window_onload
 if (seachform.nq.value <> "") and (tablerow2.style.display="none") then
   containstext.style.display=""
   switchmodetext.innerhtml="Simple<br>mode"
   tablerow2.style.display=""
 end if
end sub
</script>

<table border=0 align=center cellspacing=0 cellpadding=3>
<tr><td align=center valign=center>
   <form name=seachform action=searchdocs.asp method=get>
   <table border=0 width=100% cellspacing=0 cellpadding=2 align=center>
   <tr>
   <td valign=center align=right>
     <span id=containstext style="font-weight:bold;<%=searchformt1%>">Contains:</span>
     <input name=q type=text size=40 value='<%=q%>'>
   </td>
   <td align=right>  
     <input name=submit type=submit value="Search">
   </td>  
   <td align=right valign=center width=60>
     <a href="#" class=cool id=switchmodetext><%=searchformt2%></a>
   </td>  
   </tr>
   <tr id=tablerow2 <%=searchformt3%>>
   <td valign=center align=right>
     <span style="font-weight:bold;">Not Contains:</span>
     <input name=nq type=text size=40 value='<%=nq%>'>
   </td>
   <td colspan=2>&nbsp;</td>  
   </tr>
   </table>
   </form>
</td></tr>
</table>
<%end sub%>
