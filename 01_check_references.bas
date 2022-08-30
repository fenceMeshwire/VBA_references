Option Explicit

Sub check_reference()

Dim vbProj As Object ' Current VBA project.
Dim chkRef As Object ' Reference object.

' Refer to the activedocument's VBA project.
Set vbProj = ActiveDocument.VBProject

' Check through all selected references in the "References" dialog box.
For Each chkRef In vbProj.References
   Debug.Print "Checking reference:" & chkRef.Name
   ' If the reference is broken, debug.print the reference's name.
   If chkRef.IsBroken Then Debug.Print "Broken reference: " & chkRef.Name
Next

End Sub
