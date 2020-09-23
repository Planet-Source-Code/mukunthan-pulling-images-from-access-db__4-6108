<div align="center">

## PULLING IMAGES FROM ACCESS DB


</div>

### Description

You can get the image from database and can display in the screen
 
### More Info
 
Database must have to contain the image.

Returns image and can display the corresponding record

ASP


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MUKUNTHAN](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mukunthan.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mukunthan-pulling-images-from-access-db__4-6108/archive/master.zip)





### Source Code

```
<html>
<head>
<TITLE>Getting Employee Info</TITLE>
</head>
<body>
<h1>Employee Info</h1>
<form method="post" action="GetEmpInfo.asp">
Enter the last name of the employee you want to search:
<br><input type=text name="empLastName" size=40 value=<%=Request("empLastName")%> >
<input type=submit style="width=150" value="Show">
</form>
<hr>
<span>Sample names are: <i>Davolio, Fuller, Buchanan</i></span>
<hr><br>
<%
  if Request("empLastName") = "" Then Response.End
%>
<%
  sql1 = "select firstname, lastname, city, hiredate, photo from Employees where "
  sql2 = "lastname='"
  sql3 = Request.Form("empLastName") & "'"
  sql = sql1 & sql2 & sql3
  Set oRS = Server.CreateObject("ADODB.Recordset")
  oRS.CursorLocation = 3
  oRS.Open sql, "DSN=NW"
  SetImageForDisplay oRS("photo"), "ole"
'  oRS.Open "select logo from pub_info where pub_id='0736'", "DSN=PUBS;UID=sa"
'  SetImageForDisplay oRS("logo"), "gif"
  Set oRS.ActiveConnection = Nothing
%>
<table>
<tr>
<td valign=top><b>Employee:</b><br>
<%=oRS("firstName") %>&nbsp;<%=oRS("lastName") %><br>
<b>from </b> <%=oRS("city") %><br>
<b>hired </b> <%=oRS("hiredate") %><br>
</td>
<td>
<img src="theImg.asp"</img>
</td>
</tr>
</table>
<%
Function SetImageForDisplay(field, contentType)
  OLEHEADERSIZE = 78
  contentType = LCase(contentType)
  select case contentType
    case "gif", "jpeg", "bmp"
      contentType = "image/" & contentType
      bytes = field.value
    case "ole"
      contentType = "image/bmp"
      nFieldSize = field.ActualSize
      oleHeader = field.GetChunk(OLEHEADERSIZE)
      bytes = field.GetChunk(nFieldSize - OLEHEADERSIZE)
  end select
  Session("ImageBytes") = bytes
  Session("ImageType") = contentType
End Function
%>
</body>
</html>
```

