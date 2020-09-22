VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FE&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   390
   LinkTopic       =   "Form2"
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   26
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

  Dim ret As Long

    ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, ret
    SetLayeredWindowAttributes Me.hWnd, &HFE, 200, LWA_COLORKEY Or LWA_ALPHA

End Sub

