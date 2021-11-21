Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swFeat As SldWorks.Feature
        Set swFeat = swModel.SelectionManager.GetSelectedObject6(1, -1)
        
        If Not swFeat Is Nothing Then
            
            If swFeat.GetTypeName2() = "CoordSys" Then
                
                Dim swCoordSys As SldWorks.CoordinateSystemFeatureData
                
                Set swCoordSys = swFeat.GetDefinition
                Dim swMathTransform As SldWorks.MathTransform
                Set swMathTransform = swCoordSys.Transform
                
                Dim vMatrix As Variant
                vMatrix = swMathTransform.ArrayData
                
                Debug.Print vMatrix(0) & "," & vMatrix(1) & "," & vMatrix(2)
                Debug.Print vMatrix(3) & "," & vMatrix(4) & "," & vMatrix(5)
                Debug.Print vMatrix(6) & "," & vMatrix(7) & "," & vMatrix(8)
                Debug.Print vMatrix(9) & "," & vMatrix(10) & "," & vMatrix(11) & "," & vMatrix(12)
                
            Else
                MsgBox "Selected feature is not a coordinate system"
            End If
        Else
            MsgBox "Please select coordinate system feature"
        End If
        
    Else
        MsgBox "Please open model"
    End If
    
End Sub
