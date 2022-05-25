<!--#include file="Databaseutility.asp"-->


<%


dim response

    set response = GetRecordSet("getempdata",null)
    
    dim counter
    if not response.EOF then
       
       response.Write("<table class='GeneratedTable'>")
       response.Write("<thead>")
       response.Write("<tr>")
       
       for counter = 0 to response.fields.count-1
              response.Write("<th>" & response.fields(counter).name & vbCrLf & "</th>")
             
       next
       response.Write("</tr>")
       response.Write("</thead>")
       response.Write("<tbody>")
       
       
       do until response.EOF
                response.Write("<tr>")
                response.Write("<td>" & response("empno") & vbCrLf & "</td>")
                response.Write("<td>" & response("ename") & vbCrLf & "</td>")
                response.Write("<td>" & response("job") & vbCrLf & "</td>")
                response.Write("<td>" & response("sal") & vbCrLf & "</td>")
                response.Write("<td>" & response("dname") & vbCrLf & "</td>")
                response.Write("<td>" & response("loc") & vbCrLf & "</td>")
                response.Write("</tr>")
                
       response.moveNext
       loop
       Response.Write("</tr>") 
       response.Write("</tbody>")
       response.Write("</table>")
        
       response.write("</br>") 
       response.Write("<form method='post' action='AddNewRecord.asp'>" _
                        + "<input type='submit' value='AddNewRecord'> </form>")
    end if
%>

<html>
<title>Employee Details</title>
<style>
    table.GeneratedTable {
        width: 100%;
        background-color: #ffffff;
        border-collapse: collapse;
        border-width: 2px;
        border-color: #ffcc00;
        border-style: solid;
        color: #000000;
    }

        table.GeneratedTable td, table.GeneratedTable th {
            border-width: 2px;
            border-color: #ffcc00;
            border-style: solid;
            padding: 3px;
        }

        table.GeneratedTable thead {
            background-color: #ffcc00;
        }
</style>
</html>