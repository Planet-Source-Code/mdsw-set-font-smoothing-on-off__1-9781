VERSION 5.00
Begin VB.Form frmfontsmooth 
   AutoRedraw      =   -1  'True
   Caption         =   "#"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   Icon            =   "frmfontsmooth.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "#"
      Height          =   465
      Left            =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2625
   End
   Begin VB.PictureBox picFS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   60
      ScaleHeight     =   1095
      ScaleWidth      =   5340
      TabIndex        =   5
      Top             =   870
      Width           =   5400
   End
   Begin VB.CommandButton cmdSetSmoothOff 
      Caption         =   "#"
      Height          =   465
      Left            =   2850
      TabIndex        =   4
      Top             =   3045
      Width           =   2625
   End
   Begin VB.CommandButton cmdSetSmoothOn 
      Caption         =   "#"
      Height          =   465
      Left            =   45
      TabIndex        =   1
      Top             =   3045
      Width           =   2625
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "#"
      Height          =   465
      Left            =   2850
      TabIndex        =   0
      Top             =   3600
      Width           =   2625
   End
   Begin VB.Label lblDum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2685
      TabIndex        =   8
      Top             =   2010
      Width           =   150
   End
   Begin VB.Label lblDum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2685
      TabIndex        =   6
      Top             =   2430
      Width           =   150
   End
   Begin VB.Label lblDum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2685
      TabIndex        =   3
      Top             =   150
      Width           =   150
   End
   Begin VB.Label lblDum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2685
      TabIndex        =   2
      Top             =   540
      Width           =   165
   End
End
Attribute VB_Name = "frmfontsmooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------Set Font smoothing On/Off----------------------'
'                                                                   '
'Written By, Kenneth Marshall                                       '
'Copyright © 2000 by MoonDance Software All rights reserved.        '
'MoonDance Software © 1996 - 2000                                   '
'P.O. Box 2376                                                      '
'Kennesaw, GA 30144                                                 '
'E-mail: Moooond@aol.com                                            '
'*******************************************************************'
Option Explicit
'!-----------------------------------------------------------------!'
'NOTES: Pay close attention to the Parameters passed to this Function
'!-----------------------------------------------------------------!'
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPIF_UPDATEINIFILE = &H1 'used when setting On or Off
Private Const SPIF_SENDWININICHANGE = &H2 'used when setting On or Off
Private Const SPI_GETFONTSMOOTHING = 74 'used when reading if On or Off
Private Const SPI_SETFONTSMOOTHING = 75 'used when setting On or Off
Private Const SPI_GETWINDOWSEXTENSION = 92 'used to find out if Plus! Eetensions are installed

Private bYesNoPlus As Boolean 'this is the only one we want to hold!

Private Function DisplayOnOffYesNo(OOYN As Single, bTyp As Boolean) As String
    'pass 1 and boolean for On/Off
    'pass 2 and boolean for Yes/No
    DisplayOnOffYesNo = Choose(OOYN, IIf(bTyp, LoadResString(116), LoadResString(117)), IIf(bTyp, LoadResString(118), LoadResString(119)))
End Function

Private Sub PrintSampleFont(bTF As Boolean)
    'This Sub displays our text depending upon if the function passed or failed
    picFS.Cls 'clear it to reset our X and Y points default, this will refresh the print
    picFS.FontName = "Times New Roman"
    picFS.FontItalic = bTF
    picFS.FontBold = bTF
    picFS.FontSize = IIf(bTF, 48, 16)
    picFS.ForeColor = IIf(bTF, 0, 255) 'if everything is ok then print it black otherwise red
    picFS.Print IIf(bTF, LoadResString(111), LoadResString(112))
End Sub

Private Function SetInfoTypes(sngN As Single) As String
    'If were passing anything here then we turn turn things Off and display why
    lblDum(1).ForeColor = 255
    cmdSetSmoothOn.Enabled = False
    cmdSetSmoothOff.Enabled = False
    SetInfoTypes = Choose(sngN, LoadResString(113), LoadResString(114), LoadResString(115))
End Function

Private Sub zStrings()
    lblDum(2).Caption = LoadResString(104)
    cmdSetSmoothOn.Caption = LoadResString(105)
    cmdSetSmoothOff.Caption = LoadResString(106)
    cmdAbout.Caption = LoadResString(107)
    cmdUnload.Caption = LoadResString(108)
    frmfontsmooth.Caption = LoadResString(109)
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, frmfontsmooth
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub cmdSetSmoothOn_Click()
Dim X As Long, b As Boolean
    'turn it On
    'NOTE Parameters->   1&, ByVal vbNullString
    X = SystemParametersInfo(SPI_SETFONTSMOOTHING, 1&, ByVal vbNullString, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
    
    'Test it
    If X <> 0 Then
        'read it
        'NOTE Parameters->   0&, b
        X = SystemParametersInfo(SPI_GETFONTSMOOTHING, 0&, b, 0&)
        lblDum(1) = LoadResString(102) & DisplayOnOffYesNo(1, b)
        
        PrintSampleFont bYesNoPlus 're display info
    Else
        lblDum(1) = SetInfoTypes(1) 'all else display info and turn off controls
    End If
    
End Sub

Private Sub cmdSetSmoothOff_Click()
Dim X As Long, b As Boolean
    'turn it Off
    'NOTE Parameters->   0&, ByVal vbNullString
    X = SystemParametersInfo(SPI_SETFONTSMOOTHING, 0&, ByVal vbNullString, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
    
    'Test it
    If X <> 0 Then
        'read it
        'NOTE Parameters->   0&, b
        X = SystemParametersInfo(SPI_GETFONTSMOOTHING, 0&, b, 0&)
        lblDum(1) = LoadResString(102) & DisplayOnOffYesNo(1, b)
        
        PrintSampleFont bYesNoPlus 're display info
    Else
        lblDum(1) = SetInfoTypes(1) 'all else display info and turn off controls
    End If
    
End Sub

Private Sub Form_Load()
Dim X As Long, b As Boolean

    zStrings 'load strings

    'First order of business YOU MUST find out if Plus! is installed on the system!
    'NOTE Parameters->   1&, ByVal 0&
    X = SystemParametersInfo(SPI_GETWINDOWSEXTENSION, 1&, ByVal 0&, 0&)
    
    bYesNoPlus = CBool(X) 'convert a True or False and hold it in the program
    
    'display if Plus! is installed
    lblDum(0) = LoadResString(101) & DisplayOnOffYesNo(2, bYesNoPlus)
    
    If bYesNoPlus Then 'only if Plus! is installed then read it
    
        'pass a boolean to receive True or False
        'read it
        'NOTE Parameters->   0&, b
        X = SystemParametersInfo(SPI_GETFONTSMOOTHING, 0&, b, 0&)
        
        If X <> 0 Then
            'display if its On or Off
            lblDum(1) = LoadResString(102) & DisplayOnOffYesNo(1, b)
        Else
            lblDum(1) = SetInfoTypes(3) 'all else display info and turn off controls
        End If
    Else
        lblDum(1) = SetInfoTypes(2) 'all else display info and turn off controls
    End If
    
    PrintSampleFont bYesNoPlus 'display info
    
    lblDum(3) = LoadResString(103) & picFS.FontName 'display fontname picture box is using
End Sub


