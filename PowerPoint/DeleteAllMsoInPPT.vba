Option Explicit
Sub DeleteAllMsoInPPT()
     'This macro will only work if there is an active PowerPoint
     'It removes all mso*** in the active PowerPoint
    Dim objApp As Object, objSlide As Object, ObjShp As Object, objTable As Object
    Dim i As Long
    
    On Error Resume Next
    'Is the PowerPoint open?
    Set objApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    
    If objApp Is Nothing Then Exit Sub
    
    If objApp.ActivePresentation Is Nothing Then Exit Sub
    
    For i = 1 To 20
        Set objSlide = objApp.ActivePresentation.Slides(i)
        For Each ObjShp In objSlide.Shapes
            Select Case ObjShp.Type
                Case msoFreeform, msoLine, msoSlicer
                    ObjShp.Delete
            End Select
        Next
    Next i
End Sub
