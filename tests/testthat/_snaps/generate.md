# macro script generation

    Code
      result
    Output
      Option Explicit
      On Error Resume Next
      
      Dim args
      Set args = wScript.Arguments
      
      Dim a
      Dim b
      Dim c
      
      a = args(0)
      b = args(1)
      c = args(2)
      
      Dim macroApp
      Dim macroFile
      
      Set macroApp = createObject("Excel.Application")
      macroApp.visible = False
      Set macroFile = macroApp.Workbooks.Open("testfile.xlsx")
      
      If Err.Number <> 0 Then
        WScript.Echo "VBScript Error: " & Err.Description
        macroApp.Quit
        wscript.quit
      End If
      
      macroApp.Run "fake_macro", CInt(a), CDbl(b), CStr(c)
      
      If Err.Number <> 0 Then
        WScript.Echo "VBScript Error: " & Err.Description
        macroApp.Quit
        wscript.quit
      End If
      
      macroFile.Close
      macroApp.Quit

