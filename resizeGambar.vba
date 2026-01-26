Sub ResizeObjects()
    Dim tSlide As Slide
    Dim shp As Shape
    Dim sWidth As Single, sHeight As Single
    
    For Each tSlide In ActivePresentation.Slides
        For Each shp In tSlide.Shapes
            ' cek kalau shape adalah link ke Excel (OLE linked object)
            If shp.Type = msoLinkedOLEObject Then
                shp.LockAspectRatio = msoFalse
                ' ukuran fix (cm ? point)
                shp.Width = 22.53 * 28.346
                shp.Height = 21.33 * 28.346
                
                ' hitung posisi tengah
                sWidth = ActivePresentation.PageSetup.SlideWidth
                sHeight = ActivePresentation.PageSetup.SlideHeight
                shp.Left = (sWidth - shp.Width) / 2
                shp.Top = (sHeight - shp.Height) / 2
            End If
        Next shp
    Next tSlide
End Sub
