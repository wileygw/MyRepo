Attribute VB_Name = "MoveUp"
Sub MoveUp(R As Range, Optional col% = 0)

Dim C As Range, x%

    x = 1

    For Each C In R

        On Error GoTo ErrHandler
        
        If C.Value = "" Then
        
            Do While IsEmpty(C.Offset(x, 0)) = True
            
                If Intersect(R, C.Offset(x, 0)) Is Nothing Then Exit Sub
            
                x = x + 1
            
            Loop
            
            If col > 0 Then
                For col = 0 To col
                
                    C.Offset(0, col).Value = C.Offset(x, col).Value
            
                    C.Offset(x, col).ClearContents
                    
                Next col
                
            Else
                
                C.Offset(0, 0).Value = C.Offset(x, 0).Value
            
                C.Offset(x, 0).ClearContents
                
            End If
                    
        End If
        
    Next C
    
ErrHandler:
    Exit Sub
    
End Sub