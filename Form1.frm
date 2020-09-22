VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1065
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   810
      Left            =   7950
      ScaleHeight     =   750
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   105
      Width           =   1215
      Begin VB.VScrollBar VScroll1 
         Height          =   360
         Left            =   675
         Max             =   -1
         Min             =   -20
         TabIndex        =   7
         Top             =   375
         Value           =   -5
         Width           =   270
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Elasticity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0032CFEB&
         Height          =   300
         Left            =   30
         TabIndex        =   9
         Top             =   15
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H0032CFEB&
         Height          =   375
         Left            =   225
         TabIndex        =   8
         Top             =   345
         Width           =   450
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Random Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0032CFEB&
      Height          =   360
      Left            =   6255
      TabIndex        =   5
      Top             =   675
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Height          =   540
      Left            =   6630
      Picture         =   "Form1.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Height          =   555
      Left            =   9450
      Picture         =   "Form1.frx":34D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   405
      Width           =   780
   End
   Begin VB.PictureBox picNotifier 
      Height          =   630
      Left            =   6645
      ScaleHeight     =   570
      ScaleWidth      =   690
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1410
      Width           =   750
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   135
      Top             =   1290
   End
   Begin VB.PictureBox picSlide 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1665
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   1
      Top             =   2535
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0032CFEB&
      Height          =   705
      Left            =   135
      MaxLength       =   25
      TabIndex        =   0
      Top             =   150
      Width           =   6090
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H0032CFEB&
      Height          =   345
      Left            =   10170
      MouseIcon       =   "Form1.frx":4946
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   -75
      Width           =   300
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Change_text 
         Caption         =   "Change Text Properties"
      End
      Begin VB.Menu Start_up 
         Caption         =   "Run at Start Up"
      End
      Begin VB.Menu nothing 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
         Begin VB.Menu about_index 
            Caption         =   "             Agustin Rodriguez"
            Index           =   0
         End
         Begin VB.Menu about_index 
            Caption         =   "E-mail: virtual_guitar_1@hotmail.com"
            Index           =   1
         End
      End
      Begin VB.Menu nothing1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const NIM_ADD As Integer = &H0
Private Const NIM_MODIFY As Integer = &H1
Private Const NIM_DELETE As Integer = &H2
Private Const WM_MOUSEMOVE As Integer = &H200
Private Const NIF_MESSAGE As Integer = &H1
Private Const NIF_ICON As Integer = &H2
Private Const NIF_TIP As Integer = &H4
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LetterFollow
    X As Single
    Y As Single
    Letter As String
End Type

Private TheForm As NOTIFYICONDATA
Private PT As POINTAPI

Private xx() As New Form2
Private Amount_letter As Integer
Private Time_to_use_menu As Integer
Private Letter(-1 To 25) As LetterFollow
Private Elasticity As Integer

Private Sub Change_text_Click()

  Dim i As Integer

    For i = 0 To Amount_letter
        Unload xx(i)
    Next i
    Me.Show

End Sub

Private Sub Command1_Click()

  Dim i As Integer

    Amount_letter = Len(Text1)
    ReDim xx(Amount_letter)

    For i = 0 To Amount_letter

        With xx(i)
            .Label1(0) = Mid(Text1.Text, i + 1, 1)
            .Label1(1) = Mid(Text1.Text, i + 1, 1)
            .Show
            .Label1(1).ForeColor = ForeColor
            .Label1(0).Font = FontName
            .Label1(1).Font = FontName
            .Label1(0).FontSize = FontSize
            .Label1(1).FontSize = FontSize
            .Label1(0).FontItalic = FontItalic
            .Label1(1).FontItalic = FontItalic
            .Label1(0).FontBold = FontBold
            .Label1(1).FontBold = FontBold

            If Check1 Then
                .Label1(1).ForeColor = Int(Rnd * &HFFFFFF)
              Else
                .Label1(1).ForeColor = ForeColor
            End If
        End With

        If i Then
            xx(i - 1).Width = TextWidth(Mid(Text1, i, 1)) * Screen.TwipsPerPixelX
            xx(i - 1).Height = TextHeight(Mid(Text1, i, 1)) * Screen.TwipsPerPixelY
        End If

    Next i

    Hide
    Timer2.Enabled = True

End Sub

Private Sub Command2_Click()

  Dim Ret_font As SelectedFont

    Ret_font = ShowFont(Me.hWnd, FontName, True)
    If Ret_font.bCanceled Then
        Exit Sub
    End If
    With Ret_font
        FontName = .sSelectedFont
        FontSize = .nSize
        ForeColor = .lColor
        FontItalic = .bItalic
        FontBold = .bBold
    End With

End Sub

Private Sub Exit_Click()

  Dim i As Integer

    For i = 0 To Amount_letter
        Unload xx(i)
    Next i

    Unload Me

    End

End Sub

Private Sub Form_Load()

  Dim i As Integer
  Dim v As Long

    Read_Ini

    Randomize
    ' See if the program is set to run at startup.
    m_IgnoreEvents = True
    If WillRunAtStartup(App.EXEName) Then
        Start_up.Checked = True
      Else
        Start_up.Checked = False
    End If
    m_IgnoreEvents = False

    Command1_Click

    Put_on_Systray

    Elasticity = 5

    v = Int(Rnd * &HFFFFFF)

    FormGradient Me, 1, 1, 1, 197, 192, 158

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TheForm.cbSize = Len(TheForm)
    TheForm.hWnd = picNotifier.hWnd
    TheForm.uId = 1&
    ' Remove it from the TaskBar.
    Shell_NotifyIcon NIM_DELETE, TheForm

    Save_Ini

End Sub

Private Sub Label3_Click()

    Exit_Click

End Sub

Private Sub picNotifier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Static Rec As Boolean, Msg As Long

    Msg = X / Screen.TwipsPerPixelX
    If Rec = False Then
        Rec = True
        Select Case Msg
          Case WM_LBUTTONDBLCLK

          Case WM_LBUTTONDOWN

            Time_to_use_menu = 500
            PopupMenu Menu

          Case WM_LBUTTONUP

          Case WM_RBUTTONDBLCLK

          Case WM_RBUTTONDOWN

          Case WM_RBUTTONUP

            PopupMenu Menu

        End Select

        Rec = False
    End If

End Sub

Private Sub Timer2_Timer()

  Dim i As Integer
  Dim t As Integer
  Static on_top_timer As Single

    If Visible Then
        Exit Sub
    End If

    If on_top_timer < Timer And Visible = False Then
        If Time_to_use_menu = 0 Then
            For i = 0 To Amount_letter
                apiSetWindowPos xx(i).hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
            Next i
            on_top_timer = Timer + 1
          Else
            Time_to_use_menu = Time_to_use_menu - 1

        End If
    End If

    GetCursorPos PT

    Letter(-1).X = PT.X
    Letter(-1).Y = PT.Y

    DoEvents

    For i = 0 To Len(Text1.Text)

        If i Then
            Letter(i).X = Letter(i).X + (Letter(i - 1).X + xx(i - 1).ScaleWidth - Letter(i).X) / Elasticity
          Else
            Letter(i).X = Letter(i).X + (Letter(i - 1).X + xx(i).ScaleWidth - Letter(i).X) / Elasticity
        End If

        Letter(i).Y = Letter(i).Y + (Letter(i - 1).Y - Letter(i).Y) / Elasticity
        xx(i).Move Letter(i).X * Screen.TwipsPerPixelX, Letter(i).Y * Screen.TwipsPerPixelY

    Next i

    t = t + 1

End Sub

Private Sub Put_on_Systray()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    TheForm.cbSize = Len(TheForm)
    TheForm.hWnd = picNotifier.hWnd
    TheForm.uId = 1&
    TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TheForm.ucallbackMessage = WM_MOUSEMOVE
    TheForm.hIcon = Me.Icon
    TheForm.szTip = "Cursor Message" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, TheForm
    Me.Hide
    App.TaskVisible = False

End Sub

Private Sub VScroll1_Change()

    Label1 = Abs(VScroll1)
    Elasticity = Label1

End Sub

Private Sub Start_up_Click()

    If m_IgnoreEvents Then
        Exit Sub
    End If

    Start_up.Checked = Start_up.Checked Xor -1

    SetRunAtStartup App.EXEName, App.Path, Abs((Start_up.Checked))

End Sub

Public Sub FormGradient(TheForm As Form, RedStart, GreenStart, BlueStart, RedEnd, GreenEnd, BlueEnd)

  Dim i As Integer, j As Integer, Y As Single, H As Integer
  Dim Rk As Single, Gk As Single, Bk As Single
  Dim R As Integer, G As Integer, b As Integer
  Dim Params() As Variant
  Dim ctlObj As Control
  Dim ContObj As Control
  Dim yScale As Single

    Rk = (RedStart - RedEnd) / 1024
    Gk = (GreenStart - GreenEnd) / 1024
    Bk = (BlueStart - BlueEnd) / 1024

    On Error Resume Next

        With TheForm
            .AutoRedraw = True
            .DrawStyle = vbInsideSolid
            .DrawMode = vbCopyPen
            .ScaleMode = vbPixels
            .DrawWidth = 2
            .ScaleHeight = 1024
        End With

        For Y = 0 To 1023
            j = Y
            R = RedStart - j * Rk
            G = GreenStart - j * Gk
            b = BlueStart - j * Bk
            TheForm.Line (0, Y)-(Screen.Width, Y - 1), RGB(R, G, b), B
        Next Y

        i = 0
        ReDim Params(TheForm.Count, 5)
        For Each ctlObj In TheForm
            Params(i, 0) = LCase(TypeName(ctlObj))
            Params(i, 1) = LCase(ctlObj.Name)
            Params(i, 2) = LCase(ctlObj.Container.Name)
            Params(i, 3) = CInt(ctlObj.Top)
            Params(i, 4) = CInt(ctlObj.Height)

            If Params(i, 0) = LCase("Label") Then
                ctlObj.BackStyle = 0
              Else
                Y = Params(i, 3)
                H = Params(i, 4)
                Y = Y + H / 2
                If Params(i, 2) = LCase(TheForm.Name) Then
                    Params(i, 5) = Y
                End If
            End If
            i = i + 1
        Next ctlObj

        i = 0
        For Each ctlObj In TheForm

            If Params(i, 1) <> LCase(TheForm.Name) Then
                yScale = TheForm.ScaleHeight / TheForm.Height
                For j = 0 To TheForm.Count
                    If (j <> i) And (Params(j, 1) = Params(i, 2)) Then
                        Params(i, 5) = Params(j, 5)
                        j = TheForm.Count
                    End If
                Next j
            End If
            i = i + 1
        Next ctlObj

        i = 0
        For Each ctlObj In TheForm
            If Params(i, 5) > 0 Then
                Y = Params(i, 5)
                j = Y
                R = RedStart - j * Rk
                G = GreenStart - j * Gk
                b = BlueStart - j * Bk
                ctlObj.BackColor = RGB(R, G, b)
            End If
            i = i + 1
        Next ctlObj
    On Error GoTo 0

End Sub

Private Sub Save_Ini()

  Dim objBag            As New PropertyBag
  Dim vntBagContents    As Variant

    With objBag
        .WriteProperty "Elasticity", VScroll1.Value
        .WriteProperty "Random_Color", Check1.Value
        .WriteProperty "Text", Text1.Text
        .WriteProperty "Font", Font
        .WriteProperty "Forecolor", ForeColor
        vntBagContents = .Contents
    End With

    Open App.Path & "\Ini.Bag" For Binary As 1
    Put #1, , vntBagContents
    Close 1

End Sub

Private Sub Read_Ini()

  Dim objBag            As New PropertyBag
  Dim vntBagContents    As Variant

    If Dir(App.Path & "\Ini.Bag") = "" Then
        Font.Size = 26
        Text1.Text = "Messenger Cursor"
        Check1.Value = 1
        VScroll1.Value = -7
        Exit Sub
    End If

    Open App.Path & "\Ini.Bag" For Binary As 1
    
    Get #1, , vntBagContents
    Close 1

    With objBag

        .Contents = vntBagContents

        VScroll1.Value = .ReadProperty("Elasticity")
        Check1.Value = .ReadProperty("Random_color")
        ForeColor = .ReadProperty("Forecolor")
        Set Font = .ReadProperty("Font")
        Text1.Text = .ReadProperty("Text")

    End With

End Sub

