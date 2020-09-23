VERSION 5.00
Begin VB.Form fScreensaver 
   BorderStyle     =   0  'None
   Caption         =   "Exploding Flowers Screesaver"
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   735
   FillColor       =   &H8000000F&
   ForeColor       =   &H80000008&
   Icon            =   "fScreensaver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   38
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   49
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "fScreensaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ORIGINAL CONCEPT BY:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' EXPLODING FLOWERS (for screen saver)
''' By Paul Bahlawan
''' March 8, 2004
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Fleshed out somewhat by:
' -----------------------------------
' redbird77@earthlink.net (2004.04.18)

Option Explicit

Private hBrushNew As Long
Private hBrushOld As Long
Private oDIB      As cDIB32
Private bData()   As RGBQUAD
Private SA        As SAFEARRAY2D
Private m_bActive As Boolean
Private m_lH      As Long
Private m_lW      As Long
Private tF()      As udtFlower

Public Sub DoScreensaver()

Dim tDev As DEVMODE
Dim lRet As Long
Dim i    As Long
Dim s    As String
Dim lRes(3, 1) As Long

    Randomize
    
    ' Horrible redundant coding here.  Will change later, but for now I wanted to
    ' get the screensaver working without spending all my time on the screen
    ' resolution bit.
    lRes(0, 0) = 640: lRes(0, 1) = 480
    lRes(1, 0) = 800: lRes(1, 1) = 600
    lRes(2, 0) = 1024: lRes(2, 1) = 768
    lRes(3, 0) = 1280: lRes(3, 1) = 1024
    
    ' Do not change the screen resolution or hide cursor if only a preview.
    If tSet.Mode = "Exploding Flowers Screensaver" Then

        ' Enumerate display settings. (Taken from MSDN).
        'tDev.dmSize = LenB(tDev)

        Do
            lRet = EnumDisplaySettings(0&, i, tDev)
            's = s & tDev.dmPelsWidth & " x " & tDev.dmPelsHeight & " : " & tDev.dmBitsPerPel & " : " & tDev.dmDisplayFlags & vbCrLf
            i = i + 1
        Loop While lRet

        ' Change user's screen resolution.
        tDev.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        tDev.dmPelsWidth = lRes(tSet.Buffer.ScreenIndex, 0)
        tDev.dmPelsHeight = lRes(tSet.Buffer.ScreenIndex, 1)

        lRet = ChangeDisplaySettings(tDev, 0&)

        ShowCursor False
    End If
    
'    MsgBox "Changing screen resloution " & vbCrLf & _
           "from: " & vbCrLf & _
           "to: " & GetDeviceCaps(Me.hdc, HORZRES) & " x " & GetDeviceCaps(Me.hdc, VERTRES)
    
     ' Create a DIB as a buffer for all the drawing.
    Set oDIB = New cDIB32
    
    ' Width and height are stored as percentages of screen resolution.
    m_lW = tSet.Buffer.Width / 100 * GetDeviceCaps(Me.hdc, HORZRES)
    m_lH = tSet.Buffer.Height / 100 * GetDeviceCaps(Me.hdc, VERTRES)
    
    oDIB.Create m_lW, m_lH, tSet.Buffer.BackColor
    
    ' Set the framerate text color.
    SetTextColor oDIB.hDIBDC, &HFF
   
    ' Point the array at the DIB's bits.
    pvBuildSA SA, oDIB
    CopyMemory ByVal VarPtrArray(bData()), VarPtr(SA), 4
    
    ' Set FillColor.
    hBrushNew = CreateSolidBrush(tSet.Buffer.FillColor)
    hBrushOld = SelectObject(oDIB.hDIBDC, hBrushNew)
    Debug.Assert hBrushOld

    ReDim tF(tSet.Flower.FlowerCount - 1)
    
    m_bActive = True
    Show
    DoFlowers

End Sub

' Each petal can reuse the flowers points array.
' Petals do not need separate height, width, color, etc, since they all
' are the same (except for the points which do not need to be saved).

Public Sub DoFlowers()

Dim f       As Integer
Dim p       As Integer
Dim iPad    As Integer
Dim zAngCur As Single
Dim zAngRad As Single
Dim tmp     As Long
Dim hPenNew As Long
Dim hPenOld As Long
Dim lRet    As Long
Dim tFR     As udtFrameRate

    ' Quick fix to hide the un-blurred border that comes with the quick blur method.
    If tSet.Blur.Quick Then iPad = 10

    Do While m_bActive

        ' Draw the framerate.
        If tSet.Buffer.DisplayFrameRate Then TextOut oDIB.hDIBDC, 5, 5, tFR.Text, Len(tFR.Text)
    
        ' Create the flowers.
        f = Int(Rnd * 100)
    
        If f <= UBound(tF) Then
    
        If tF(f).PetalCount = 0 Then
            
            ' Example on how the RndEx function works.
            ' -----------------------------------------------------
            ' PetalCount setting can be from 1(sparse) - 10(bushy).
            
            ' PetalCount is set to 1 then RndEx returns (3-5).
            ' PetalCount is set to 2 then RndEx returns (4-6).
            ' PetalCount is set to 10 then RndEx returns (12-14).
            
            ' Increase the v(ariance) parameter to return a wider
            ' range of numbers centered around the n parameter.
            ' -----------------------------------------------------
            tF(f).PetalCount = RndEx(tSet.Flower.PetalCount + 3, 2)
        
            tF(f).Center.x = Int(Rnd() * m_lW)
            tF(f).Center.y = Int(Rnd() * m_lH)
        
            Do
                tF(f).Direction.x = Int(Rnd() * 9) - 4
                tF(f).Direction.y = Int(Rnd() * 9) - 4
            Loop Until tF(f).Direction.x And tF(f).Direction.y
        
            ' Flower angle and spin.
            tF(f).Angle = Int(Rnd * 360)
            tF(f).Spin = Int(Rnd() * 7) - 3
            
            ' Petal width and height.
            tF(f).PetalWidth = Int(Rnd() * 25) + 1
            'tF(f).PetalWidth = Int(Rnd() * (360 / tF(f).PetalCount)) + 1
            tF(f).PetalHeight = Int(Rnd() * (m_lH \ 2)) + 20
            
            tF(f).Pointiness = Int(Rnd() * (tF(f).PetalHeight * 0.19 * _
                               (10 - tSet.Flower.PetalPointiness)))
            
            tF(f).Bounce = Rnd()
            tF(f).BounceRate = Rnd() / 20
            
            ' Returns 2 to 254.
            tF(f).Color.r.Value = Int(Rnd() * 253) + 2
            tF(f).Color.g.Value = Int(Rnd() * 253) + 2
            tF(f).Color.b.Value = Int(Rnd() * 253) + 2
    
            ' Returns -2 to 2.
            tF(f).Color.r.Direction = Int(Rnd() * 5) - 2
            tF(f).Color.g.Direction = Int(Rnd() * 5) - 2
            tF(f).Color.b.Direction = Int(Rnd() * 5) - 2
        
        End If
        
    End If
    
    ' Draw the flowers.
    For f = 0 To UBound(tF)

        ' Create a new outline color for each flower at each position.
        hPenNew = CreatePen(0, 1, RGB(tF(f).Color.r.Value, _
                                      tF(f).Color.g.Value, _
                                      tF(f).Color.b.Value))
                                      
        hPenOld = SelectObject(oDIB.hDIBDC, hPenNew): Debug.Assert hPenOld
    
        ' ..petal by petal.
        For p = 0 To tF(f).PetalCount - 1

            ' Current angle = start angle + petal position offset.
            zAngCur = tF(f).Angle + (360 / tF(f).PetalCount * p)

            ' Set current petal point 1 of 4.
            tF(f).Points(0).x = tF(f).Center.x
            tF(f).Points(0).y = tF(f).Center.y
            
            ' Set current petal point 2 of 4.
            zAngRad = (zAngCur + tF(f).PetalWidth) * PiDiv180
            tmp = tF(f).Pointiness * tF(f).Bounce
            tF(f).Points(1).x = tF(f).Center.x + tmp * Cos(zAngRad)
            tF(f).Points(1).y = tF(f).Center.y + tmp * Sin(zAngRad)

            ' Set current petal point 3 of 4.
            zAngRad = zAngCur * PiDiv180
            tmp = tF(f).PetalHeight * tF(f).Bounce
            tF(f).Points(2).x = tF(f).Center.x + tmp * Cos(zAngRad)
            tF(f).Points(2).y = tF(f).Center.y + tmp * Sin(zAngRad)
            
            ' Set current petal point 4 of 4.
            zAngRad = (zAngCur - tF(f).PetalWidth) * PiDiv180
            tmp = tF(f).Pointiness * tF(f).Bounce
            tF(f).Points(3).x = tF(f).Center.x + tmp * Cos(zAngRad)
            tF(f).Points(3).y = tF(f).Center.y + tmp * Sin(zAngRad)

            lRet = BeginPath(oDIB.hDIBDC): Debug.Assert lRet
            
            Polygon oDIB.hDIBDC, tF(f).Points(0), 4
            
            lRet = EndPath(oDIB.hDIBDC): Debug.Assert lRet
            
            ' If FillColor is "transparent" then just stroke path
            ' otherwise stroke and fill path.
            If tSet.Buffer.FillColor = -1 Then
                lRet = StrokePath(oDIB.hDIBDC)
            Else
                lRet = StrokeAndFillPath(oDIB.hDIBDC)
            End If
            
        Next ' Petal
        
        ' Clean up the pen.
        lRet = SelectObject(oDIB.hDIBDC, hPenOld): Debug.Assert lRet
        lRet = DeleteObject(hPenNew): Debug.Assert lRet
        
        ' Move flowers.
        With tF(f).Center
            .x = .x + tF(f).Direction.x
            .y = .y + tF(f).Direction.y
        End With
        
        ' Rotate and bounce the flowers.
        With tF(f)
        
            .Angle = .Angle + .Spin
            
            .Bounce = .Bounce + .BounceRate
            If .Bounce > 1 Or .Bounce < 0.05 Then .BounceRate = -.BounceRate
            
        End With
        
        ' Set colors.
        tF(f).Color.r.Value = tF(f).Color.r.Value + tF(f).Color.r.Direction
        If tF(f).Color.r.Value > 254 Or tF(f).Color.r.Value < 2 Then
            tF(f).Color.r.Direction = -tF(f).Color.r.Direction
        End If
        
        tF(f).Color.g.Value = tF(f).Color.g.Value + tF(f).Color.g.Direction
        If tF(f).Color.g.Value > 254 Or tF(f).Color.g.Value < 2 Then
            tF(f).Color.g.Direction = -tF(f).Color.g.Direction
        End If
        
        tF(f).Color.b.Value = tF(f).Color.b.Value + tF(f).Color.b.Direction
        If tF(f).Color.b.Value > 254 Or tF(f).Color.b.Value < 2 Then
            tF(f).Color.b.Direction = -tF(f).Color.b.Direction
        End If

        ' Terminate flowers.
        If tF(f).Center.x < -tF(f).PetalHeight * tF(f).Bounce Or _
           tF(f).Center.x > m_lW + tF(f).PetalHeight * tF(f).Bounce Then
                tF(f).PetalCount = 0
        End If
        
        If tF(f).Center.y < -tF(f).PetalHeight * tF(f).Bounce Or _
           tF(f).Center.y > m_lH + tF(f).PetalHeight * tF(f).Bounce Then
                tF(f).PetalCount = 0
        End If
            
    Next ' Flower
    
        If tSet.Blur.Enabled Then
            If tSet.Blur.Quick Then
                Blur oDIB.Width - 1, oDIB.Height - 1
            Else
                BlurCustom oDIB.Width - 1, oDIB.Height - 1
            End If
        End If
        
        ' Stretch the buffer over the screen.
        oDIB.Stretch Me.hdc, _
                     -iPad, -iPad, _
                     Me.ScaleWidth + iPad + iPad, Me.ScaleHeight + iPad + iPad, _
                     0, 0, _
                     m_lW, m_lH, tSet.Buffer.StretchMode
    
        ' Calculate framerate.
        If tSet.Buffer.DisplayFrameRate Then
           With tFR
               .Value = .Value + 1
               If (timeGetTime - .Ticks >= 1000) Then
                   .Text = .Value & " fps"
                   .Ticks = timeGetTime
                   .Value = 0
               End If
           End With
        End If
        
    DoEvents
    
    If Not m_bActive Then Exit Sub
    
Loop

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
Static xPos As Integer
Static yPos As Integer
    
    If xPos = 0 Then
        xPos = x
        yPos = y
    Else
        If Abs(xPos - x) > 5 And (tSet.Mode = "Exploding Flowers Screensaver") Then Unload Me: Exit Sub
        If Abs(yPos - y) > 5 And (tSet.Mode = "Exploding Flowers Screensaver") Then Unload Me: Exit Sub
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If tSet.Mode = "Exploding Flowers Screensaver" Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim r As Long

    ' Clean up.
    m_bActive = False
    
    r = SelectObject(oDIB.hDIBDC, hBrushOld): Debug.Assert r
    r = DeleteObject(hBrushNew): Debug.Assert r
    
    CopyMemory ByVal VarPtrArray(bData), 0&, 4
    Set oDIB = Nothing
    
    If tSet.Mode = "Exploding Flowers Screensaver" Then ShowCursor True
    
    ' Revert to the orginal screen resolution.
    r = ChangeDisplaySettings(ByVal 0&, 0)
    
    PutSettings
    
End Sub

Public Sub Blur(ByVal cx As Long, ByVal cy As Long)

' Red, green, and blue accumulators.
Dim r As Long, g As Long, b As Long
Dim x As Long, y As Long

    For x = 1 To cx - 1
        For y = 1 To cy - 1
        
            ' Reset.
            r = 0: g = 0: b = 0
            
            ' 1 4 7
            ' 2 5 8
            ' 3 6 9
            
            r = r + bData(x - 1, y - 1).r
            g = g + bData(x - 1, y - 1).g
            b = b + bData(x - 1, y - 1).b

            r = r + bData(x - 1, y).r
            g = g + bData(x - 1, y).g
            b = b + bData(x - 1, y).b

            r = r + bData(x - 1, y + 1).r
            g = g + bData(x - 1, y + 1).g
            b = b + bData(x - 1, y + 1).b

            r = r + bData(x, y - 1).r
            g = g + bData(x, y - 1).g
            b = b + bData(x, y - 1).b

            r = r + bData(x, y).r
            g = g + bData(x, y).g
            b = b + bData(x, y).b

            r = r + bData(x, y + 1).r
            g = g + bData(x, y + 1).g
            b = b + bData(x, y + 1).b

            r = r + bData(x + 1, y - 1).r
            g = g + bData(x + 1, y - 1).g
            b = b + bData(x + 1, y - 1).b

            r = r + bData(x + 1, y).r
            g = g + bData(x + 1, y).g
            b = b + bData(x + 1, y).b

            r = r + bData(x + 1, y + 1).r
            g = g + bData(x + 1, y + 1).g
            b = b + bData(x + 1, y + 1).b
           
            ' Divide and set.
            bData(x, y).r = r \ 9
            bData(x, y).g = g \ 9
            bData(x, y).b = b \ 9
            
        Next
    Next

End Sub

Public Sub BlurCustom(ByVal cx As Long, ByVal cy As Long)

' Red, green, and blue accumulators.
Dim r As Long, g As Long, b As Long
Dim x As Long, y As Long, i As Long, j As Long
Dim iOffset As Integer, vn As Integer

    iOffset = tSet.Blur.Strength

    For x = 0 To cx
        For y = 0 To cy
        
            ' Reset.
            r = 0: g = 0: b = 0: vn = 0
            
            For i = -iOffset To iOffset
                For j = -iOffset To iOffset
                
                    If (x + i > 0 And x + i <= cx) And _
                       (y + j > 0 And y + j <= cy) Then
            
                    r = r + bData(x + i, y + j).r
                    g = g + bData(x + i, y + j).g
                    b = b + bData(x + i, y + j).b
                    
                    vn = vn + 1
                    
                    End If
                
                Next
            Next

            ' Divide and set.
            bData(x, y).r = r \ vn
            bData(x, y).g = g \ vn
            bData(x, y).b = b \ vn
            
        Next
    Next

End Sub

Private Sub pvBuildSA(ByRef tSA As SAFEARRAY2D, ByRef DIB As cDIB32)
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.Width
        .pvData = DIB.DIBitsPtr
    End With
End Sub

' -------------------------------------------------------------------
' Helper Functions
' -------------------------------------------------------------------
Private Function RndEx(ByVal n As Long, ByVal v As Integer) As Long

Dim hi As Long, lo As Long

    hi = n + v: lo = n - v
    RndEx = Int((hi - lo + 1) * Rnd + lo)

End Function

Private Function RndEx2(ByVal n As Long, ByVal v As Single) As Long
' Generates a random number +/-v% of n.
    RndEx2 = Interpolate(n * (1 - v), n * (1 + v), Rnd)
End Function

Private Function Interpolate(ByVal a As Long, ByVal b As Long, ByVal p As Single) As Long
    Interpolate = a * (1 - p) + b * p
End Function
