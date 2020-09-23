VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEnDe1 
   Caption         =   "Code\Decode"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   645
      Top             =   420
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Code 1 (Simple) Unchecked Code2 (More Complex) Checked"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   0
      Width           =   2625
   End
   Begin VB.CommandButton cmdDecode 
      BackColor       =   &H0080C0FF&
      Caption         =   "Decode"
      Height          =   345
      Left            =   3075
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4830
      Width           =   915
   End
   Begin VB.CommandButton cmdCode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Code"
      Height          =   345
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2775
      Width           =   915
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   795
      TabIndex        =   11
      Top             =   45
      Width           =   3330
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Load"
      Height          =   330
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   795
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   75
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4905
      MaxLength       =   6
      TabIndex        =   6
      Top             =   465
      Width           =   1170
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Save"
      Height          =   330
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear"
      Height          =   330
      Left            =   1635
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   795
   End
   Begin VB.TextBox txtDecoded 
      Height          =   1680
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5190
      Width           =   7335
   End
   Begin VB.TextBox txtEncoded 
      Height          =   1680
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3135
      Width           =   7335
   End
   Begin VB.TextBox txtSource 
      Height          =   1680
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label Label5 
      Caption         =   "Filename"
      Height          =   240
      Left            =   135
      TabIndex        =   12
      Top             =   75
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   "DeCoded text"
      Height          =   240
      Left            =   105
      TabIndex        =   9
      Top             =   4950
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Coded text"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2925
      Width           =   1470
   End
   Begin VB.Label Label4 
      Caption         =   "Text to be Coded"
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   870
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "      Key Value (1 - 255)"
      Height          =   300
      Left            =   4635
      TabIndex        =   5
      Top             =   855
      Width           =   2130
   End
End
Attribute VB_Name = "frmEnDe1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'set textbox margins
Private Const EM_SETMARGINS = &HD3
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Public Sub RightMargin(hWnd As Long, n As Integer)
    SendMessageLong hWnd, EM_SETMARGINS, EC_RIGHTMARGIN, n * &H10000
End Sub

Public Sub LeftMargin(hWnd As Long, n As Integer)
    SendMessageLong hWnd, EM_SETMARGINS, EC_LEFTMARGIN, n
End Sub

Public Sub SetSelected()
  ' used to auto select text in txtEncoded window
  
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub

Public Function Crypt(Source As String, strPassword As String, EnDeCrypt As Boolean) As String
    'Code 2
    'EnDeCrypt True = Encrypt
    'EnDeCrypt False = Decrypt
    Dim intPassword As Long
    Dim intCrypt As Long
    Dim x As Integer

    For x = 1 To Len(strPassword)
        intPassword = intPassword + Asc(Mid$(strPassword, x, 1))
    Next x


    For x = 1 To Len(Source)


        If EnDeCrypt = True Then
            intCrypt = Asc(Mid$(Source, x, 1)) + intPassword + x
            


            Do Until intCrypt <= 255
                intCrypt = intCrypt - 255
            Loop
        Else
            intCrypt = Asc(Mid$(Source, x, 1)) - intPassword - x
            


            Do Until intCrypt > 0
                intCrypt = intCrypt + 255
            Loop
        End If
        Crypt = Crypt & Chr(intCrypt)
    Next x
End Function

Public Sub LoadText(TBox As TextBox, file As String)

   On Error GoTo error
   Dim mystr As String
   Dim x As String
   Dim T$
   Dim texas$

   Open file For Input As #1
   Do While Not EOF(1)
            Line Input #1, T$
            texas$ = texas$ + T$ + Chr$(13) + Chr$(10)
        Loop
        TBox = texas$
        Close #1
   Exit Sub

error:
  x = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

Public Sub SaveText(TBox As TextBox, file As String)

   On Error GoTo error
   Dim mystr As String
   Dim x As String

   Open file For Output As #1
   Print #1, TBox
   Close 1
   Exit Sub

error:
  x = MsgBox("There has been a error!", vbOKOnly, "Error")
End Sub

Function EnDeCode(strOld As String, bcode As Byte) As String
  'Code 1
  'This bit of code is by Joshua Larsen from his
  'programmed called EnDeCode
  
    Dim strChar As String ' temp char holder
    
    While Len(strOld) > 0 ' until entire line is encoded
        strChar = Left(strOld, 1) ' get single char
        strOld = Right(strOld, Len(strOld) - 1) ' remove char from old string
        strChar = Chr$(bcode Xor Asc(strChar)) ' encode char
        EnDeCode = EnDeCode + strChar ' add char to new string
        
    Wend
    
End Function

Private Sub cmdClear_Click()
  txtSource.Text = ""
  txtEncoded.Text = ""
  txtDecoded.Text = ""
  txtFilename.Text = ""
  txtSource.SetFocus
End Sub

Private Sub cmdCode_Click()
If Check1.Value = 0 Then
        txtEncoded.Text = EnDeCode(txtSource.Text, Int(Val(txtKey.Text)))
   Else
        txtEncoded.Text = Crypt(txtSource, txtKey, True)
   End If
   
   txtEncoded.SetFocus
End Sub

Private Sub cmdDecode_Click()
If Check1.Value = 0 Then
   txtDecoded.Text = EnDeCode(txtEncoded.Text, Int(Val(txtKey.Text)))
Else
   txtDecoded.Text = Crypt(txtEncoded, txtKey, False)
End If
End Sub

Private Sub cmdLoad_Click()
   CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
   CommonDialog1.ShowOpen
   txtFilename.Text = CommonDialog1.FileName
   If txtFilename.Text = "" Then Exit Sub
   Call LoadText(txtEncoded, txtFilename.Text)
End Sub

Private Sub cmdSave_Click()

   CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
   CommonDialog1.ShowSave
   txtFilename.Text = CommonDialog1.FileName
   If txtFilename.Text = "" Then Exit Sub
   Call SaveText(txtEncoded, txtFilename.Text)
End Sub

Private Sub Form_Load()
  'Set margins in textboxes
  
    LeftMargin txtSource.hWnd, 5
    RightMargin txtSource.hWnd, 5
    LeftMargin txtEncoded.hWnd, 5
    RightMargin txtEncoded.hWnd, 5
    LeftMargin txtDecoded.hWnd, 5
    RightMargin txtDecoded.hWnd, 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me

End Sub

Private Sub Timer1_Timer()
If Check1.Value = 0 Then
   Label1.Caption = "      Key Value (1 - 255)"
 Else
   Label1.Caption = "Password 1 to 6 Charactors"
 End If
End Sub
Private Sub txtEncoded_GotFocus()
SetSelected  'auto select all text in txtEncoded window

End Sub


