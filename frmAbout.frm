VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "http://members.xoom.com/marshall48/MiscVB/vbaddins.htm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      MouseIcon       =   "frmAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1380
      Width           =   5940
   End
   Begin VB.Label lblDum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1050
      Width           =   3600
   End
   Begin VB.Label lblDum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   1
      Left            =   3390
      TabIndex        =   1
      Top             =   645
      Width           =   150
   End
   Begin VB.Label lblDum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   75
      Width           =   210
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub zStrings()
frmAbout.Caption = LoadResString(107)
lblDum(0).Caption = LoadResString(109)
lblDum(1).Caption = LoadResString(110)
End Sub

'Open default browser to web site
Public Sub OpenWebsite(strWebsite As String)
Dim opWst As Long
opWst = ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL)
    
    If opWst < 33 Then
        Select Case opWst
            Case 2, 3
                '----------------------------------------------
                MsgBox "The Path:" & vbLf & _
                       "'" & strWebsite & "'" & vbLf & _
                       "Was Not Found.", 48, "Link Not Found"
                '----------------------------------------------
            Case 31
                '----------------------------------------------
                MsgBox "*********************************************************************" & vbLf & _
                       "!     No Association Found for HTTP Addresses.                !" & vbLf & _
                       "!     A Web Browser should be Installed or reinstalled         !" & vbLf & _
                       "*********************************************************************", 48, "No Association In Registry"
                '----------------------------------------------
            Case Is <= 32
                '----------------------------------------------
                MsgBox "An error occurred attempting to open your browser to path..." & vbLf & _
                       "'" & strWebsite & "'" & vbLf & _
                       "ERROR Return code: ( " & opWst & " )", 16, "ERROR"
                '----------------------------------------------
            Case Else
                'move on
        End Select
    End If

End Sub

Private Sub Form_Load()
    zStrings
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not lblLink.ForeColor = &HFF0000 Then lblLink.ForeColor = &HFF0000
End Sub


Private Sub lblDum_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not lblLink.ForeColor = &HFF0000 Then lblLink.ForeColor = &HFF0000
End Sub


Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not lblLink.ForeColor = &HFF& Then lblLink.ForeColor = &HFF&
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OpenWebsite("http://members.xoom.com/marshall48/MiscVB/vbaddins.htm")
End Sub

