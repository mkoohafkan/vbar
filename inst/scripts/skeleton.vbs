Option Explicit
On Error Resume Next

Dim args
Set args = wScript.Arguments

{dim_args}

{assign_args}

Dim macroApp
Dim macroFile

Set macroApp = createObject("{macro_application}.Application")
macroApp.visible = False
Set macroFile = macroApp.{collection_type}.Open("{macro_file}")

If Err.Number <> 0 Then
  WScript.Echo "VBScript Error: " & Err.Description
  macroApp.Quit
  wscript.quit
End If

macroApp.Run "{macro_name}", {macro_args}


If Err.Number <> 0 Then
  WScript.Echo "VBScript Error: " & Err.Description
  macroApp.Quit
  wscript.quit
End If

macroFile.Close
macroApp.Quit
