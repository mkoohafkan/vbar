Option Explicit
On Error Resume Next

Dim args
Set args = wScript.Arguments

{dim_args}

{assign_args}

Dim macroApp
Dim macroFile

set macroApp = createObject("{macro_application}.Application")
macroApp.visible = False
set macroFile = macroApp.{collection_type}.Open("{macro_file}")

macroApp.Run "{macro_name}", {macro_args}

macroFile.Close
macroApp.Quit
