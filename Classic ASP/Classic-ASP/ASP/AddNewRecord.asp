<!--#include file="Databaseutility.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
 

<title>Candy</title>

</head>

<body>
<form name=frmAddCandy action="AddNewRecord.asp" method=post>

<p>EmpName: <input type="text" name="txtEmpName" size="20"></p>
<p>Job:<input type="text" name="txtjob" size="20"></p> 
<p>Salary:<input type="text" name="txtsalary" size="20"></p> 



<p><input type="submit" value="Add Employee" name="butAdd">&nbsp;




</form>
</body>

</html>



<%
    
    dim empname, job, salary
    empname =  request.form("txtEmpName")
    job =  request.form("txtjob")
    salary =  request.form("txtsalary")

    'Checking if user lands first time, should not call this functionality except user submit form
    If not IsEmpty(empname) and empname<>"" Then

        dim p_empname , p_job , p_salary
        'Preparing parameterized command 
        set p_empname = CreateParameters("@ename", adVarChar, adParamInput , 400 , empname)
        set p_job = CreateParameters("@job", adVarChar, adParamInput , 400 , job)
        set p_salary = CreateParameters("@salary", adDouble, adParamInput , 400 , salary)

        dim paramCollection : paramCollection = Array( p_empname , p_job, p_salary )
        dim error_code , isdataSave

        'Executing 
        call ExecuteStoreProcedure("saveemp" , paramCollection , null , 0)    

        
        response.write("<script language=""javascript"">alert ('Data save sucessfully!'); window.location='DisplayEmployeeData.asp'</script>")

        
    End If

    


    


%>