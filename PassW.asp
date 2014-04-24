<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
function makePassword(byVal maxLen) 

Dim strNewPass 
Dim whatsNext, upper, lower, intCounter 
Randomize 

For intCounter = 1 To maxLen 
 whatsNext = Int((1 - 0 + 2) * Rnd + 0) 
If whatsNext = 0 Then 
'character 
 upper = 90 
 lower = 65 
Else 
 if whatsNext=1 then
  upper=122
  lower=97
 else
  upper = 57 
  lower = 48 
 end if
End If 
 strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower)) 
Next 
 makePassword = strNewPass 
end function 

Dim jj
%>
<style type="text/css">
<!--
.STYLE2 {
	font-weight: bold;
	font-size: xx-large;
}
-->
</style>

<table width="95%" border="1" align="center" cellpadding="2" cellspacing="7" bordercolor="#FFCC00" bordercolordark="#FFFFFF" bgcolor="#f8ffe8">
  <caption>
  <span class="STYLE2">901Ëæ»úµÇÂ½Âë</span>
  </caption>
  <tr>
    <% 
for jj=1 to ubound(c_allstat)+1 
%>
    <td width="10%"><%=c_allstat(jj-1)%></td>
    <td width="12%"><%=makePassword(8)%></td>
    <% If jj mod 4=0 Then response.Write("</tr><tr>") end if 
next%>
  </tr>
</table>
