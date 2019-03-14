<% @Language = "VBScript" %>
<% Option Explicit %>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>IIS Excel Interop Test</title>
    <link href="style.css" rel="stylesheet">
  </head>
  <body>
  <h1>IIS Excel Interop Test</h1>
  <h3>This page will test if excel can be used from within classic asp code under IIS.</h3>
  <div class="hint">
  <p>When running Microsoft Excel using COM automation you need to take care of some special treatment.</p>
  <ul>
    <li>
  <p>
  First of all you have to make sure that there is absolutely now user input required at any time. 
  </p>
  <p>
  As soon as excel requests any user input (i.e. file overwrite, save as, vb runtime error) your iis process will hang.<br />
  This will result in a timeout error on the client side and a hanging/orphan excel process on the server.<br />
  <b>Error handling in visual basic script code is essential</b>
  </p>
    </li>
    <li>
  <p>
  Second, when running excel using COM automation via Server.CreateObject(...) there will no additional startup files be loaded.<br />
  You have to load all additional files mnaually on your own.
  </p>
    </li>
  </ul>
  <p>
  Now let us see if we can start excel and do some workbook operation.<br />You can download the workbook used for testing <a href="Book1.xlsm">here</a>.
  </p>
  </div>
  <div class="paper">
<%
' This is a simple VBScript to test IIS Excel Interop.
'  
' The test launches Excel, loads a workbook with vb script code.
' Then it runs a macro within excel.
' The macro sets Sheet1!A1 to "Test OK"
' After this the asp script read workBook.Worksheets(1).Range("A1")
' and shows its value.
'

CONST CRLF = "<br />"

Sub Log(s)
  Dim dtNow, sTimestamp
  dtNow = Now
  'sTimestamp = Right(Year(dtNow), 2) & Right("00" & Month(dtNow), 2) & Right("00" & Day(dtNow), 2)
  sTimestamp = sTimestamp & " " & Right("00" & Hour(dtNow), 2) & ":" & Right("00" & Minute(dtNow), 2) & ":" & Right("00" & Second(dtNow), 2)

  Response.Write sTimestamp & " : " & s & CRLF
End Sub
   
   Dim objWSHNetwork, xlApp, fileName, workBook
   
   Set objWSHNetwork = Server.CreateObject("WScript.Network")
  
   Log "Testing Excel IIS Interop on <b>" & objWSHNetwork.ComputerName & "</b>"
    
   Log "Server.CreateObject( 'Excel.Application' )..."
   Set xlApp = Server.CreateObject( "Excel.Application" )
   Log "   ...OK"

           
   With xlApp

     Log "Excel.LibraryPath = " & .LibraryPath

     .Visible = False
     .AlertBeforeOverwriting = False
     .AskToUpdateLinks = False
     .DisplayAlerts = False
     .Interactive = False
     .ScreenUpdating = False

     fileName = Server.MapPath(".\Book1.xlsm")
     Log ".Workbooks.Open('" & fileName & "')..."
     Set workBook = .Workbooks.Open(fileName)
     Log "   ...OK"


     Log ".Run 'TestExcelInterop'..."
     .Run "TestExcelInterop"
     Log "Result is <b>" & workBook.Worksheets(1).Range("A1").Value & "</b>"
     Log "   ...OK"

 
     Log ".Workbooks.Close..."
     .Workbooks.Close
     Log "   ...OK"

     Log ".Quit"     
      .Quit
     Log "   ...OK"

   End With
   
   Log "Done."      
   Set xlApp = Nothing   
%>
  </div>
  <div class="copy">Copyright &copy; 2019 by <a href="https://github.com/AndiSHFR/excel-interop-in-iis-with-classic-asp" />Andreas Schaefer</div>
  </body>
</html>
