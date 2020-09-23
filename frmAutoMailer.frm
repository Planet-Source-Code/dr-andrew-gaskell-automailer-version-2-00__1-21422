VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAutoMailer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoMailer"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9405
   ForeColor       =   &H00000000&
   Icon            =   "frmAutoMailer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImglstTool 
      Left            =   840
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoMailer.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoMailer.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoMailer.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoMailer.frx":077E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImglstTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold Font"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic Font"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline Font"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh status window"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtprofile 
      Height          =   285
      Left            =   6600
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtsubject 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox txtedit 
      Height          =   2535
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmAutoMailer.frx":0892
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   8055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   240
      Top             =   6120
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgMailer 
      Left            =   240
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSubject 
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblProfile 
      Caption         =   "Profile for Mail Session: (leave blank to be prompted)  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label lblBody 
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFont 
         Caption         =   "F&ont"
      End
      Begin VB.Menu mnuTextColor 
         Caption         =   "&Text Color"
      End
   End
End
Attribute VB_Name = "frmAutoMailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Public oSession As Mapi.Session
Public oMessage As Mapi.Message
Public oRecip As Mapi.Recipient
Public MessageID As String
Dim bcancel As Boolean


Private Sub cmdExit_Click()

    On Error GoTo ErrorHandler

    If cmdStart.Caption = "&Stop" Then
        cmdStart_Click
    End If

    DoEvents

    Unload Me

Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub
Private Sub cmdStart_Click()

    On Error GoTo ErrorHandler

    If cmdStart.Caption = "&Start" Then
        cmdStart.Caption = "&Stop"
        Log Now & vbTab & "---- AutoMailer services started"
        Timer1.Enabled = True
        frmAutoMailer.MousePointer = vbHourglass
    Else
        cmdStart.Caption = "&Start"
        Log Now & vbTab & "---- AutoMailer services stopped"
        Timer1.Enabled = False
    End If

Exit Sub

ErrorHandler:
    Select Case Err.Number
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub SendMessage(sRecipient As String, sMessage As String)
    'This sub is the key code in the program.
    'It uses CDO and the mapirtf.dll to send RTF formatted e-mail (no need to write html)
    'mapi controls alone won't do this
    'the new message is added to the outbox, addresing and title added, the ID of the
    'message obtained, RTF property added to the message with that ID
    'then the message sent
    'Full error handling included, in case the user cancels or enters incorrect profile
    
    Dim oMsgFilter As Mapi.MessageFilter
    Dim bRet As Integer
    Dim strProfileName As String
    Dim blogon As Boolean
    strProfileName = txtprofile.Text
    blogon = False
    On Error GoTo ErrorHandler:
    
    Set oSession = CreateObject("MAPI.Session")
    If oSession Is Nothing Then
        MsgBox "Could not create Mapi Session", vbOKOnly, "VBSendRTF"
        End
    End If

    If Not oSession Is Nothing And strProfileName <> "" Then
        oSession.Logon ProfileName:=strProfileName, _
                    showDialog:=False
        
    ElseIf Not oSession Is Nothing And strProfileName = "" Then
        frmAutoMailer.MousePointer = vbDefault
        oSession.Logon
        frmAutoMailer.MousePointer = vbHourglass
   End If
   
   blogon = True

    
    CreateNewMessage
    frmAutoMailer.MousePointer = vbHourglass
    
    
    
    'Set Subject
    oMessage.Subject = txtsubject.Text
    
    
    'Set Recipient
    Set oRecip = oMessage.Recipients.Add
    oRecip.Name = sRecipient
    oRecip.Type = ActMsgTo
    oRecip.Resolve  'check names
    

    'Save the message
    oMessage.Update
    
    'Set the RTF property, using the MAPIRTF dll
    bRet = WriteRTF(oSession.Name, oMessage.ID, _
                    oMessage.StoreID, sMessage)
    If Not bRet = 0 Then
        MsgBox "RTF Property not stored successfully!", vbOKOnly, "VBSendRTF Warning"
    End If
    
    ' Because the object has changed, we must re-create
      ' our Message object variable
      ' First clear our current variable
    Set oMessage = Nothing
    ' Next set a filter on the outbox for our Message ID
    Set oMsgFilter = oSession.Outbox.Messages.Filter
    oMsgFilter.Fields(ActMsgPR_ENTRYID) = MessageID
    Set oMessage = oSession.Outbox.Messages.GetFirst
    
    ' Clear the Message Filters
    Set oMsgFilter = Nothing
    Set oSession.Outbox.Messages.Filter = Nothing
    
    ' Send the message
    oMessage.Send
    oSession.Logoff
    Set oSession = Nothing
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        'flags set to zero, no more mail to be sent
        Case 3021
            Log Now & vbTab & "---- No more mail to be sent. Stopping Automailer Services"
            frmAutoMailer.MousePointer = vbDefault
            cmdStart_Click
            Exit Sub
        
        'user entered incorrect profile
        Case -2147221231
            Log Now & vbTab & "---- Incorrect/Unknown Profile. Stopping Automailer Services"
            frmAutoMailer.MousePointer = vbDefault
            cmdStart_Click
            Exit Sub
        
        'user pressed cancel on the choose new email address dialog box
        Case -2147221229 And blogon = True
            Log Now & vbTab & "---- User cancelled send"
            bcancel = True
            oSession.Outbox.Messages.Delete
    
            Set oSession.Outbox.Messages.Filter = Nothing
            Set oSession = Nothing
            Exit Sub
               
        'user pressed cancel when the logon dialog appeared
        Case -2147221229 And blogon = False
            Log Now & vbTab & "---- User cancelled this Logon session."
            bcancel = True

            Set oSession = Nothing
            
          'other MAPI CDO error
        Case Else
            Log Err.Number & " - " & Err.Description
            
    End Select
    
End Sub
Sub CreateNewMessage()
    'add a message to the outbox
    Set oMessage = oSession.Outbox.Messages.Add
    Log Now & vbTab & "---- New message added to Outbox"
    If oMessage Is Nothing Then
        MsgBox "Could not create Message", vbOKOnly, "VBSendRTF"
        End
    End If
    oMessage.Update 'save message
    MessageID = oMessage.ID 'get message ID
End Sub
Private Sub Form_Load()

    Dim sDBLocation As String
    Dim hSysMenu As Long

    On Error GoTo ErrorHandler

    ' disable the 'X':
    hSysMenu = GetSystemMenu(hwnd, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

    Log "AutoMailer version " & App.Major & "." & App.Minor & App.Revision
    Log "Press start to begin services"

    sDBLocation = App.Path & "\@bcusers.mdb"    'set name/location of database here

    Set cn = New ADODB.Connection
    cn.Open "PROVIDER=MSDASQL;" & _
                "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                "DBQ=" & sDBLocation & ";" & _
                "UID=;PWD=;"
                
Exit Sub

ErrorHandler:
    Select Case Err.Number
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub

Private Sub mnuExit_Click()
    On Error GoTo ErrorHandler

    If cmdStart.Caption = "&Stop" Then
        cmdStart_Click
    End If

    DoEvents

    Unload Me

Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select
End Sub

Private Sub mnuFont_Click()
    dlgMailer.Flags = cdlCFBoth
    dlgMailer.CancelError = True
    On Error Resume Next
        dlgMailer.ShowFont
        If Err = 0 Then
            txtedit.SelFontName = dlgMailer.FontName
            txtedit.SelFontSize = dlgMailer.FontSize
            txtedit.SelBold = dlgMailer.FontBold
            txtedit.SelItalic = dlgMailer.FontItalic
            txtedit.SelStrikeThru = dlgMailer.FontStrikethru
            txtedit.SelUnderline = dlgMailer.FontUnderline
        Else
            Err = 0
        End If
    On Error GoTo 0
End Sub

Private Sub mnuOpen_Click()
    Dim OpenFileName As String
    Dim OpenFileText As String
    dlgMailer.CancelError = True
    dlgMailer.Flags = 0
    dlgMailer.Filter = "Text (*.txt)|*.txt|All files (*.*)|*.*"
    On Error Resume Next
        dlgMailer.ShowOpen
        If Err = 0 Then
            OpenFileName = dlgMailer.FileName
            Open OpenFileName For Binary As 1
                OpenFileText = Space$(LOF(1))
                Get 1, , OpenFileText
            Close 1
            txtedit.Text = OpenFileText
            OpenFileText = ""
        Else
            Err = 0
        End If
    On Error GoTo 0
End Sub

Private Sub mnuSave_Click()
    Dim OpenFileName As String
    dlgMailer.CancelError = True
    dlgMailer.Flags = 0
    dlgMailer.Filter = "Text (*.txt)|*.txt|All files (*.*)|*.*"
    On Error Resume Next
        dlgMailer.ShowSave
        If Err = 0 Then
            OpenFileName = dlgMailer.FileName
            Open OpenFileName For Binary As 1
                Put 1, , txtedit.Text
            Close 1
        Else
            Err = 0
        End If
    On Error GoTo 0
End Sub

Private Sub mnuTextColor_Click()
    dlgMailer.Flags = 0
    dlgMailer.CancelError = True
    On Error Resume Next
        dlgMailer.ShowColor
        If Err = 0 Then
            txtedit.SelColor = dlgMailer.Color
        Else
            Err = 0
        End If
    On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
    'This sub was adapted from Mark Wilson's
    'an e-mail is only sent to those addresses in the database
    'with a flag of 1. I made changes to the code so that the program will stop sending
    'when EOF is reached. For a mailing list you can set members whom you want to keep
    'a record of, but not mail any more, to 0.
    'I moved the error handling to the send message sub, as this provides more powerful
    'debugging.
    
    
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open "Select * from tblemail where s_flag = 1", cn, , , adCmdUnknown
    rs.MoveFirst

    If rs.RecordCount <> 0 Then
        While rs.EOF = False
            
            
            SendMessage rs.Fields("emailaddress"), txtedit.TextRTF
            
            
            DoEvents
    
    If bcancel = False Then
        Log Now & vbTab & "---- Mail sent to:  " & rs.Fields("emailaddress")
    ElseIf bcancel = True Then
        Log Now & vbTab & "---- Mail NOT sent to: " & rs.Fields("emailaddress")
        bcancel = False
    End If
            ' *********comment the following two lines for unpopular users(!) *********
            'rs.Fields("s_flag").Value = 0
            'rs.Update
            ' *****************************************************************************************
            
            rs.MoveNext
        Wend
    Else
        rs.Close
    End If
    If rs.EOF = True Then
        Log Now & vbTab & "---- No more mail to be sent. Stopping Automailer Services"
        cmdStart_Click
        Timer1.Enabled = False
        frmAutoMailer.MousePointer = vbDefault
    End If
Exit Sub

End Sub

Private Sub Log(ByVal sText As String)

    ' this way it doesnt refresh the whole thing every time, no blinking
    With txtStatus
        .SelStart = Len(.Text)
        .SelText = sText & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
        .SelLength = 0
    End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'handle toolbar button clicks
    Select Case Button.Key
    Case "Bold"
        txtedit.SelBold = Not txtedit.SelBold
    Case "Italic"
        txtedit.SelItalic = Not txtedit.SelItalic
    Case "Underline"
        txtedit.SelUnderline = Not txtedit.SelUnderline
    Case "Refresh"
        txtStatus.Text = " "
    End Select
End Sub

Private Sub txtStatus_GotFocus()

    frmAutoMailer.SetFocus

End Sub
