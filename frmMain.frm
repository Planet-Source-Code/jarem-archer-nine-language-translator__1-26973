VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WorldLingo Machine Translator"
   ClientHeight    =   6675
   ClientLeft      =   3330
   ClientTop       =   915
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox strHTML 
      Height          =   5025
      Left            =   5610
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   450
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.TextBox txtTrans 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1920
      Width           =   5145
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   150
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "Translate!"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   5970
      Width           =   1965
   End
   Begin VB.ComboBox cmbTarget 
      Height          =   315
      ItemData        =   "frmMain.frx":3665
      Left            =   3360
      List            =   "frmMain.frx":3684
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5580
      Width           =   1965
   End
   Begin VB.ComboBox cmbSource 
      Height          =   315
      ItemData        =   "frmMain.frx":36E5
      Left            =   3360
      List            =   "frmMain.frx":3704
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5220
      Width           =   1965
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear Text"
      Height          =   285
      Left            =   4170
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   4890
      Width           =   1125
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   6390
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Words: 0"
            TextSave        =   "Words: 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6985
            Text            =   "Translation Status: Waiting"
            TextSave        =   "Translation Status: Waiting"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3630
      Width           =   5145
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Target Language:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   990
      TabIndex        =   8
      Top             =   5550
      Width           =   2385
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Language:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   990
      TabIndex        =   7
      Top             =   5190
      Width           =   2385
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Original Text:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   3300
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Translated Text:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   1590
      Width           =   2085
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbSource_Click()
Select Case cmbSource.Text
    Case "French": cmbSource.Tag = "FR"
    Case "German": cmbSource.Tag = "DE"
    Case "Spanish": cmbSource.Tag = "ES"
    Case "Japanese": cmbSource.Tag = "JA"
    Case "Korean": cmbSource.Tag = "KO"
    Case "Portuguese (Brazilian)": cmbSource.Tag = "PT"
    Case "Chinese": cmbSource.Tag = "ZH"
    Case "English": cmbSource.Tag = "EN"
    Case "Italian": cmbSource.Tag = "IT"
End Select
End Sub

Private Sub cmbTarget_Click()
Select Case cmbTarget.Text
    Case "French": cmbTarget.Tag = "FR"
    Case "German": cmbTarget.Tag = "DE"
    Case "Spanish": cmbTarget.Tag = "ES"
    Case "Japanese": cmbTarget.Tag = "JA"
    Case "Korean": cmbTarget.Tag = "KO"
    Case "Portuguese (Brazilian)": cmbTarget.Tag = "PT"
    Case "Chinese": cmbTarget.Tag = "ZH"
    Case "English": cmbTarget.Tag = "EN"
    Case "Italian": cmbTarget.Tag = "IT"
End Select
End Sub
Function OpenUrl(strURL As String)
Dim strHTML As String, lngIndex As Integer, temp() As Byte

temp = Inet.OpenUrl(strURL, 1)

For lngIndex = 0 To UBound(temp) - 1
strHTML = strHTML + Chr(temp(lngIndex))
Next lngIndex

OpenUrl = strHTML
End Function

Private Sub cmdClear_Click()
txtSource = ""
End Sub

Private Sub cmdTranslate_Click()
Dim StartTime

cmdTranslate.Enabled = False

StartTime = Timer 'Sets the time started
StatusBar.Panels(2) = "Translation Status: Requesting Information..."

'First lets post the text, source and target to the worldlingo site:
strHTML = Inet.OpenUrl("http://www.worldlingo.com/wl/Translate?wl_text=" & txtSource & "&wl_gloss=1&wl_srclang=" & cmbSource.Tag & "&wl_trglang=" & cmbTarget.Tag)
On Error GoTo err

txtTrans = ParseTrans 'Parse the html
Call txtSource.SetFocus
cmdTranslate.Enabled = True
StatusBar.Panels(2) = "Translation Status: Done!"

Exit Sub
err:
MsgBox err.Description
StatusBar.Panels(2) = "Translation Status: Error"

End Sub
Function ParseTrans()
Dim StartAt As Integer, EndAt As Integer
'Lets look for the translation and set its starting posision
StartAt = InStr(1, strHTML, "<textarea name=""wl_result""")
StartAt = StartAt + 115 'add 115 to make sure it start from the begining of the translation
If StartAt = 0 Then GoTo err 'Make sure it isnt 0


'Now lets see where the translation ends:
EndAt = InStr(StartAt, strHTML, "</textarea>") - 2


'add it to the textbox
ParseTrans = Mid(strHTML, StartAt, EndAt - StartAt)

Exit Function
err: ' MsgBox "Requested Information Invalid", vbCritical
StatusBar.Panels(2) = "Translation Status: Requested Information Invalid"
End Function
Private Sub Form_Load()
cmbSource.ListIndex = 7
cmbTarget.ListIndex = 2
End Sub
Public Function HTTPSafeString(Text As String) As String
   'By: Greg Tyndall
    Dim lCounter As Long
    Dim sBuffer As String
    Dim sReturn As String
    sReturn = Text


    For lCounter = 1 To Len(Text)
        sBuffer = Mid(Text, lCounter, 1)


        If Not sBuffer Like "[a-z,A-Z,0-9]" Then
            sReturn = Replace(sReturn, sBuffer, "%" & Hex(Asc(sBuffer)))
        End If
    Next lCounter
    HTTPSafeString = sReturn
End Function
Private Sub txtSource_Change()
If Trim(txtSource) = "" Then
    StatusBar.Panels(1).Text = "Words: 0"
    Exit Sub
Else
    Dim temp
    temp = Trim(txtSource) & " " 'add a space, in case 1 word
    temp = Split(temp, " ")
    StatusBar.Panels(1).Text = "Words: " & UBound(temp)
End If

End Sub
