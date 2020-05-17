'---------------------------

'набор макросов
'ResizeBitmap – меняет разрешение всех растров
'OpenClosePath - находит открытые пути и помечает
'RGB2CMYK - переводит все растры в CMYK (сначала в rgb, чтобы в cmyk через профиль)
'PageToPowerClip - вставляет в p.clip выбранное, либо всё на стр.
'MoveToDesktop - помещает выбранные объекты на Рабочий стол
'Text2Curves - открытый макрос для конвертирования текста в кривые

'kirukir@ya.ru -------------

Sub ResizeBitmap()
    Dim p As Page
    Dim pLast As Page
    Dim res As Double
    
    res = InputBox("введите разрешение", "меняем разрешение растров")
    
    ActiveDocument.BeginCommandGroup "Replace Colors"
    On Error GoTo ErrHandler
    Set pLast = ActivePage
    For Each p In ActiveDocument.Pages
        p.Activate
        DoResizeBitmap p.Shapes, res
    Next p
    pLast.Activate

ExitSub:
    ActiveDocument.EndCommandGroup
    Exit Sub
ErrHandler:
    MsgBox "Unexpected error occured: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Sub DoResizeBitmap(ss As Shapes, ByVal res As Double)
    Dim s As Shape
    For Each s In ss
        Select Case s.Type
            Case cdrBitmapShape
                s.Bitmap.Resample , , True, res, res
            Case cdrGroupShape
                DoResizeBitmap s.Shapes, res
        End Select
        On Error Resume Next
        If Not s.PowerClip Is Nothing Then DoResizeBitmap s.PowerClip.Shapes, res
    Next s
End Sub


Sub OpenClosePath()
    Dim p As Page
    Dim pLast As Page
    Dim inf As Integer
    Dim arr() As Double
    Dim pagenum As Long
    Dim lr As Long
         
    'ActiveDocument.BeginCommandGroup "Check"
    On Error GoTo ErrHandler
    Set pLast = ActivePage
    For Each p In ActiveDocument.Pages
        p.Activate
        pagenum = p.Index
        OpenPathInfo p.Shapes, inf, arr(), pagenum
    Next p
    
    If inf > 0 Then
        For i = 0 To inf - 1
            x = arr(0, i)
            y = arr(1, i)
            pagenum = arr(2, i)
            lr = arr(3, i)
            Set s = ActiveDocument.Pages(pagenum).Layers(lr).CreateEllipse2(x, y, 0.15)
            s.Fill.UniformColor.CMYKAssign 0, 0, 100, 0
        Next i
    End If
    
    pLast.Activate
    
   
    MsgBox "в документе найдено " & inf & " открытых кривых"

ExitSub:
    'ActiveDocument.EndCommandGroup
    Exit Sub
ErrHandler:
    MsgBox "Unexpected error occured: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Sub OpenPathInfo(ss As Shapes, ii As Integer, arr_2() As Double, pn As Long)
    Dim s As Shape
    Dim x As Double, y As Double
    Dim spath As SubPath
    
    For Each s In ss
        Select Case s.Type
            Case cdrCurveShape
                If s.Curve.Closed = False Then
                    For Each spath In s.Curve.SubPaths
                        If spath.Closed = False Then
                            spath.Nodes.First.GetPosition x, y
                            ReDim Preserve arr_2(3, ii)
                            arr_2(0, ii) = x
                            arr_2(1, ii) = y
                            arr_2(2, ii) = pn
                            arr_2(3, ii) = s.Layer.Index
                            ii = ii + 1
                        End If
                    Next spath
                End If
            Case cdrEllipseShape
                If s.Ellipse.Type = cdrArc Then MsgBox "Внимание! На странице " & pn & " обнаружен открытый элипс!", vbCritical, "Ellipse found"
            Case cdrGroupShape
                OpenPathInfo s.Shapes, ii, arr_2, pn
        End Select
        On Error Resume Next
        If Not s.PowerClip Is Nothing Then OpenPathInfo s.PowerClip.Shapes, ii, arr_2, pn
    Next s
End Sub

Sub RGB2CMYK()
    Dim p As Page
    Dim pLast As Page
    
    ActiveDocument.BeginCommandGroup "Replace Colors"
    On Error GoTo ErrHandler
    Set pLast = ActivePage
    For Each p In ActiveDocument.Pages
        p.Activate
        DoRGB2CMYK p.Shapes
    Next p
    pLast.Activate

ExitSub:
    ActiveDocument.EndCommandGroup
    Exit Sub
ErrHandler:
    MsgBox "Unexpected error occured: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Sub DoRGB2CMYK(ss As Shapes)
    Dim s As Shape
    For Each s In ss
        Select Case s.Type
            Case cdrBitmapShape
                s.Bitmap.ConvertTo cdrRGBColorImage
                s.Bitmap.ConvertTo cdrCMYKColorImage
            Case cdrGroupShape
                DoRGB2CMYK s.Shapes
        End Select
        On Error Resume Next
        If Not s.PowerClip Is Nothing Then DoRGB2CMYK s.PowerClip.Shapes
    Next s
End Sub

Sub PageToPowerClip()

    Dim pclip As Shape
    Dim bleed As Double
    Dim pgSR As New ShapeRange
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    If ActiveLayer.IsSpecialLayer Then
        MsgBox "Выбран служебный слой! (направляющие, сетка, рабочий стол)"
        Exit Sub
    End If
    
    ActiveDocument.BeginCommandGroup "pc"
    
    If Not ActiveShape Is Nothing Then
        pgSR.Add ActiveSelection
    Else
        pgSR.AddRange ActivePage.Shapes.All
        For i = 1 To ActiveDocument.Pages(0).Layers.Count
            With ActiveDocument.Pages(0).Layers(i)
               If Not .Printable Then pgSR.RemoveRange .Shapes.All
            End With
        Next
    End If
    
    bleed = Val(InputBox("insert bleed", "bleed size", 0))
     
    With ActivePage
        Set pclip = ActiveLayer.CreateRectangle _
       (.CenterX - .SizeWidth / 2 - bleed, bleed + .CenterY + .SizeHeight / 2, .CenterX + .SizeWidth / 2 + bleed, .CenterY - .SizeHeight / 2 - bleed)
    End With
    
    pgSR.AddToPowerClip pclip
    
    ActiveShape.Outline.Width = 0
    
    ActiveDocument.EndCommandGroup

End Sub


Sub MoveToDesktop()
    Dim sr As New ShapeRange
    Set lr = ActiveDocument.MasterPage.DesktopLayer
    sr.Add ActiveSelection
    sr.MoveToLayer lr
End Sub

Sub Text2Curves()
    Dim p As Page
    Dim pLast As Page
    Dim res As Double
    
    ActiveDocument.BeginCommandGroup "T2C"
    On Error GoTo ErrHandler
    Set pLast = ActivePage
    For Each p In ActiveDocument.Pages
        p.Activate
        DoText2Curves p.Shapes
    Next p
    pLast.Activate
    
    MsgBox ("done!")

ExitSub:
    ActiveDocument.EndCommandGroup
    Exit Sub
ErrHandler:
    MsgBox "Unexpected error occured: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Sub DoText2Curves(ss As Shapes)
    Dim s As Shape
    For Each s In ss
        Select Case s.Type
            Case cdrTextShape
                s.ConvertToCurves
            Case cdrGroupShape
                DoText2Curves s.Shapes
        End Select
        On Error Resume Next
        If Not s.PowerClip Is Nothing Then DoText2Curves s.PowerClip.Shapes
    Next s
End Sub
