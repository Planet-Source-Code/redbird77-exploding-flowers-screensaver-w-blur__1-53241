VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exploding Flowers Settings"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Credits and Thanks"
      Height          =   1455
      Left            =   120
      TabIndex        =   32
      Top             =   3360
      Width           =   3135
      Begin VB.Label lblCap 
         Caption         =   "Carles P.V. - uber nifty cDIB32 class."
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblCap 
         Caption         =   "Paul Bahlawan - original concept and flower geometry."
         Height          =   495
         Index           =   17
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   2895
      End
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   36
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame fraBuffer 
      Caption         =   "Buffer Settings"
      Height          =   2655
      Left            =   3360
      TabIndex        =   20
      Top             =   1680
      Width           =   2895
      Begin VB.ComboBox ddlRes 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkFrameRate 
         Caption         =   "Display Frame Rate"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "Use Halftone Stretch (slow)"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox chkFillColor 
         Caption         =   "Fill Color"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   600
         TabIndex        =   23
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblFillColor 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblCap 
         Caption         =   "Back Color"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Screen Resolution"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCap 
         Caption         =   "Width           %  x  Height            %"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Frame fraBlur 
      Height          =   1575
      Left            =   3360
      TabIndex        =   14
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtBlur 
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Top             =   1080
         Width           =   375
      End
      Begin VB.OptionButton optBlur 
         Caption         =   "Custom Blur"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optBlur 
         Caption         =   "Quick Blur"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkBlur 
         Caption         =   "Blur"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblCap 
         Caption         =   "Strength             (> 2 = very slow)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   35
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame fraFlower 
      Caption         =   "Flower Settings"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.HScrollBar hsbFlower 
         Height          =   255
         Index           =   3
         LargeChange     =   2
         Left            =   720
         Max             =   10
         Min             =   1
         TabIndex        =   13
         Top             =   2280
         Value           =   1
         Width           =   1695
      End
      Begin VB.HScrollBar hsbFlower 
         Height          =   255
         Index           =   1
         LargeChange     =   2
         Left            =   720
         Max             =   10
         Min             =   1
         TabIndex        =   2
         Top             =   1440
         Value           =   1
         Width           =   1695
      End
      Begin VB.HScrollBar hsbFlower 
         Height          =   255
         Index           =   0
         LargeChange     =   2
         Left            =   720
         Max             =   20
         Min             =   3
         TabIndex        =   1
         Top             =   600
         Value           =   3
         Width           =   1695
      End
      Begin VB.Label lblCap 
         Caption         =   "... more to come?  Ideas, anyone?"
         ForeColor       =   &H80000011&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         Caption         =   "Petal Pointiness"
         Height          =   255
         Index           =   11
         Left            =   960
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Pointy"
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   8
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Blunt"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         Caption         =   "Petal Count"
         Height          =   255
         Index           =   16
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblCap 
         Caption         =   "Bushy"
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Sparse"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         Caption         =   "Flower Count"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCap 
         Caption         =   "Many"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Few"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "fSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' redbird77@earthlink.net (2004.04.16)

Private Sub chkBlur_Click()

Dim bEnabled As Boolean

    bEnabled = CBool(chkBlur.Value)
    
    optBlur(0).Enabled = bEnabled
    optBlur(1).Enabled = bEnabled
    txtBlur.Enabled = bEnabled
    lblCap(0).Enabled = bEnabled
    
End Sub

Private Sub Form_Load()

' Set the form's controls' properties based on the values of tSet (which
' was initially populated from the ini file).

    With ddlRes
        .AddItem "640 x 480"
        .AddItem "800 x 600"
        .AddItem "1024 x 768"
        .AddItem "1280 x 1024"
    End With

    With tSet.Buffer
        txtWidth.Text = .Width
        txtHeight.Text = .Height
        lblBackColor.BackColor = .BackColor
        chkFillColor.Value = IIf(.FillColor = -1, 0, 1)
        lblFillColor.BackColor = IIf(.FillColor = -1, 0, .FillColor)
        chkMode.Value = .StretchMode
        chkFrameRate.Value = .DisplayFrameRate
        ddlRes.ListIndex = .ScreenIndex
    End With

    With tSet.Blur
        chkBlur.Value = .Enabled
        chkBlur_Click
        optBlur(0).Value = .Quick
        optBlur(1).Value = Not optBlur(0).Value
        txtBlur.Text = .Strength
    End With

    With tSet.Flower
        hsbFlower(0).Value = .FlowerCount
        hsbFlower(1).Value = .PetalCount
        hsbFlower(3).Value = .PetalPointiness
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

' Save the current control values into tSet, which upon exiting the app,
' will be written to the ini file.

    With tSet.Buffer
        .Width = CLng(txtWidth.Text)
        .Height = CLng(txtHeight.Text)
        .BackColor = lblBackColor.BackColor
        .FillColor = IIf(chkFillColor.Value = 0, -1, lblFillColor.BackColor)
        .StretchMode = CLng(chkMode.Value)
        .DisplayFrameRate = CLng(chkFrameRate.Value)
        .ScreenIndex = CLng(ddlRes.ListIndex)
    End With

    With tSet.Blur
        .Enabled = CLng(chkBlur.Value)
        .Quick = CLng(optBlur(0).Value * -1)
        .Strength = CLng(txtBlur.Text)
    End With

    With tSet.Flower
        .FlowerCount = hsbFlower(0).Value
        .PetalCount = hsbFlower(1).Value
        .PetalPointiness = hsbFlower(3).Value
    End With
    
    PutSettings
    Unload Me
                  
End Sub

Private Sub lblBackColor_Click()

On Error GoTo ErrHandler

    cdlColor.ShowColor
    lblBackColor.BackColor = cdlColor.Color
    
    Exit Sub
    
ErrHandler:
    
End Sub

Private Sub lblFillColor_Click()

On Error GoTo ErrHandler

    cdlColor.ShowColor
    lblFillColor.BackColor = cdlColor.Color
    
    Exit Sub
    
ErrHandler:
   
End Sub
