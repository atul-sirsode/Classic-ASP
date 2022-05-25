# Classic-ASP
In this Project, i will demonstrate working with Classic ASP
Database Operations
File Upload
Rest API call
Form Validation
Use of Third Party Library Eg: Chilkat



# Enable Debugging in Classic ASP.
1) Open appwiz.cpl from Run window ( Ctrl + R ).
2) Turn windows features on or off
3) check all components of IIS
4) open inetmgr ( internet information services : IIS )
5) click on application pool , create new application pool -> select .NET CLR version as No Managed Code, pipeline as Classic.
6) click on server  > click on ASP component under IIS > under compilation Debugging properties -> Enable Server-side Debugging as True, and Send Errors to Browser as True.
7) Apply Changes.

# Open Visual Studio in admin-mode
1) Enable debugging from visual studio, make sure vs should open in admin-mode.
2) Hit debugger in code, want to debug
3) goto debug menu, attach to process - > in attach to : select script -> refresh - > select w3wp.exe file process -> click attached.
