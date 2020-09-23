Attribute VB_Name = "mMain"
Option Explicit

Private Type udtScreensaverCmd
    Letter As String
    hWnd   As Long
End Type

' tSet is an instance user-defined type that holds all the user settings
' and is "global" in scope.
Public tSet As udtSettings

Public Sub Main()

Dim tCmd   As udtScreensaverCmd
Dim lRet   As Long
Dim lStyle As Long
Dim lhWnd  As Long
    
    GetSettings
    
    ParseCommand Command$(), tCmd

    Select Case tCmd.Letter
    
        Case "c"
            tSet.Mode = "Exploding Flowers Settings"
            
            ' TODO: Need to show modally against hwnd of /c:hwnd.
            fSettings.Show vbModal
            
        Case "s"
            tSet.Mode = "Exploding Flowers Screensaver"
    
            ' If a previous instance of the screensaver in "screensaver"
            ' mode is running, then bail.
            If FindWindow(vbNullString, tSet.Mode) Then Exit Sub
            
            fScreensaver.Caption = tSet.Mode
            fScreensaver.DoScreensaver
            
        Case "p"
            tSet.Mode = "Exploding Flowers Preview"
            fScreensaver.Caption = tSet.Mode
            lStyle = GetWindowLong(fScreensaver.hWnd, GWL_STYLE)
            lRet = SetWindowLong(fScreensaver.hWnd, GWL_STYLE, lStyle Or WS_CHILD)

            If lRet Then
                lRet = SetParent(fScreensaver.hWnd, tCmd.hWnd)
                fScreensaver.DoScreensaver
            End If

            'lStyle = GetWindowLong(fScreensaver.hWnd, GWL_STYLE)
            'lRet = SetWindowLong(fScreensaver.hWnd, GWL_STYLE, lStyle And (Not WS_CHILD))
        
        Case ""
            tSet.Mode = "Exploding Flowers Settings"
            fSettings.Show vbModeless
            
    End Select
    
End Sub

Private Sub ParseCommand(ByVal sCommand As String, ByRef tCmd As udtScreensaverCmd)

' Possible incoming command-lines:
' /p HWND
' /S
' /c:HWND
' /a HWND
' (empty)
    
    If sCommand = "" Then Exit Sub
    
    tCmd.Letter = LCase$(Mid$(sCommand, 2, 1))
    
    If Mid$(sCommand, 3, 1) = "" Then Exit Sub
    
    tCmd.hWnd = CLng(Mid$(sCommand, 4))
    
End Sub

Public Sub GetSettings()

' Get settings from the ini file.

Dim f As String, s As String

    f = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\") & "Exploding Flowers.ini"
    
    s = "BufferSettings"
    With tSet.Buffer
        .Width = CLng(GetValue(s, "Width", f, "50"))
        .Height = CLng(GetValue(s, "Height", f, "50"))
        .BackColor = CLng(GetValue(s, "BackColor", f, "0"))
        .FillColor = CLng(GetValue(s, "FillColor", f, "-1"))
        .StretchMode = CLng(GetValue(s, "StretchMode", f, "0"))
        .DisplayFrameRate = CLng(GetValue(s, "DisplayFrameRate", f, "0"))
        .ScreenIndex = CLng(GetValue(s, "ScreenIndex", f, "1"))
    End With
    
    s = "BlurSettings"
    With tSet.Blur
        .Enabled = CLng(GetValue(s, "Enabled", f, "1"))
        .Quick = CLng(GetValue(s, "Quick", f, "1"))
        .Strength = CLng(GetValue(s, "Strength", f, "1"))
    End With
    
    s = "FlowerSettings"
    With tSet.Flower
        .FlowerCount = CLng(GetValue(s, "FlowerCount", f, "7"))
        .PetalCount = CLng(GetValue(s, "PetalCount", f, "5"))
        .PetalPointiness = CLng(GetValue(s, "PetalPointiness", f, "5"))
    End With
    
End Sub

Public Sub PutSettings()

' Save settings to the ini file.

Dim f As String, s As String

    f = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\") & "Exploding Flowers.ini"
    
    s = "BufferSettings"
    With tSet.Buffer
        PutValue s, "Width", CStr(.Width), f
        PutValue s, "Height", CStr(.Height), f
        PutValue s, "BackColor", CStr(.BackColor), f
        PutValue s, "FillColor", CStr(.FillColor), f
        PutValue s, "StretchMode", CStr(.StretchMode), f
        PutValue s, "DisplayFrameRate", CStr(.DisplayFrameRate), f
        PutValue s, "ScreenIndex", CStr(.ScreenIndex), f
    End With
        
    s = "BlurSettings"
    With tSet.Blur
        PutValue s, "Enabled", CStr(.Enabled), f
        PutValue s, "Quick", CStr(.Quick), f
        PutValue s, "Strength", CStr(.Strength), f
    End With
    
    s = "FlowerSettings"
    With tSet.Flower
        PutValue s, "FlowerCount", CStr(.FlowerCount), f
        PutValue s, "PetalCount", CStr(.PetalCount), f
        PutValue s, "PetalPointiness", CStr(.PetalPointiness), f
    End With

End Sub
