Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swFeat As SldWorks.Feature
        Set swFeat = swModel.SelectionManager.GetSelectedObject6(1, -1)
        
        If Not swFeat Is Nothing Then
            
            '选中的坐标系
            If swFeat.GetTypeName2() = "CoordSys" Then
            
                Dim swEntity     As SldWorks.Entity
                Dim swParentComp As SldWorks.Component2
                Dim swMathTransform As SldWorks.MathTransform
                

                Dim swCoordSys As SldWorks.CoordinateSystemFeatureData
                Set swCoordSys = swFeat.GetDefinition
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Transform Property (ICoordinateSystemFeatureData)
                ' https://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icoordinatesystemfeaturedata~transform.html
                '
                ' 这个变换是坐标系对象相对所在部件坐标系的变换
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '得到坐标系相对部件的变换
                Set swMathTransform = swCoordSys.transform
                
                '将 feature 转换为 entity
                Set swEntity = swFeat
                '坐标系所在的部件
                Set swParentComp = swEntity.GetComponent
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' IMultiply Method (IMathTransform)
                ' https://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imathtransform~imultiply.html
                '
                ' the result of transforming math transform with respect to the transformIn coordinate frame
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Transform2 Property (IComponent2)
                ' https://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.icomponent2~transform2.html
                '
                ' The transform is still with respect to the root component of the active assembly document
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '部件相对与根部件的变换
                Set swMathTransform = swMathTransform.Multiply(swParentComp.Transform2)
                
                '输出结果
                Dim vMatrix As Variant
                vMatrix = swMathTransform.ArrayData
                
                '输出旋转矩阵
                Debug.Print swFeat.Name & ": 旋转矩阵"
                Debug.Print vMatrix(0) & "," & vMatrix(1) & "," & vMatrix(2)
                Debug.Print vMatrix(3) & "," & vMatrix(4) & "," & vMatrix(5)
                Debug.Print vMatrix(6) & "," & vMatrix(7) & "," & vMatrix(8)
                Debug.Print vMatrix(9) & "," & vMatrix(10) & "," & vMatrix(11) & "," & vMatrix(12)
                Debug.Print
                
                '输出四元数
                Dim vQuaternion As Variant
                vQuaternion = Quaternion(vMatrix)
                
                Debug.Print swFeat.Name & ": 四元数"
                Debug.Print vQuaternion(0) & "," & vQuaternion(1) & "," & vQuaternion(2) & "," & vQuaternion(3)
                Debug.Print vMatrix(9) & "," & vMatrix(10) & "," & vMatrix(11) & "," & vMatrix(12)
                Debug.Print
                
                '输出RPY角
                Dim vEuler As Variant
                vEuler = Euler(vMatrix)
                Debug.Print swFeat.Name & ": RPY角"
                Debug.Print vEuler(0) & "," & vEuler(1) & "," & vEuler(2)
                Debug.Print vMatrix(9) & "," & vMatrix(10) & "," & vMatrix(11) & "," & vMatrix(12)
                Debug.Print
                
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

''''''''''''''''''''''''''
' 旋转矩阵计算四元数
''''''''''''''''''''''''''
Function Quaternion(m)
    Dim q(4) As Variant
    q(0) = Sqr((m(0) + m(4) + m(8) + 1)) / 2
    q(1) = Sgn(m(5) - m(7)) * Sqr(1 + m(0) - m(4) - m(8)) / 2
    q(2) = Sgn(m(6) - m(2)) * Sqr(1 - m(0) + m(4) - m(8)) / 2
    q(3) = Sgn(m(1) - m(3)) * Sqr(1 - m(0) - m(4) + m(8)) / 2
    Quaternion = q
End Function

'''''''''''''''''''''''''''
' 旋转矩阵计算RPY角
'''''''''''''''''''''''''''
Function Euler(m)
    Dim e(3) As Variant
    Const PI As Double = 3.14159265359
    If Abs(m(2)) <> 1 Then
        e(1) = -Arcsin(m(2))
        e(0) = ArcTan2(m(8) / Cos(e(1)), m(5) / Cos(e(1)))
        e(2) = ArcTan2(m(0) / Cos(e(1)), m(1) / Cos(e(1)))
    Else
        e(2) = 0
        If m(6) = -1 Then
            e(1) = PI / 2
            e(0) = ArcTan2(m(6), m(3))
        Else
            e(1) = -PI / 2
            e(0) = ArcTan2(m(6), -m(3))
        End If
        
    End If
    e(0) = e(0) * 180 / PI
    e(1) = e(1) * 180 / PI
    e(2) = e(2) * 180 / PI
    Euler = e
End Function

'''''''''''''''''''''''''
' 反正弦函数
'''''''''''''''''''''''''
Function Arcsin(X)
    Arcsin = Atn(X / Sqr(-X * X + 1))
End Function

'''''''''''''''''''''''''
' 反正切函数
'''''''''''''''''''''''''
Function ArcTan2(X, Y)
    Const PI As Double = 3.14159265359
    Select Case X
        Case Is > 0
            ArcTan2 = Atn(Y / X)
        Case Is < 0
            ArcTan2 = Atn(Y / X) + PI * Sgn(Y)
            If Y = 0 Then ArcTan2 = ArcTan2 + PI
        Case Is = 0
            ArcTan2 = PI / 2 * Sgn(Y)
    End Select
End Function
