VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Notification System"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
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
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Custom Notification"
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   5295
      Begin VB.TextBox txtSoundFileLocation 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   4935
      End
      Begin VB.CheckBox chkAllowSameKeys 
         Caption         =   "Remove Notifications with same Key"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "CUSTOM"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create Notification"
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CheckBox chkEnableClick 
         Caption         =   "Enable Clickable Link"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   4935
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Option2"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "This is an Example of Description Text"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Text            =   "Example Title"
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound File Location:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Key"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Icon"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   315
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   720
         Picture         =   "frmMain.frx":0000
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   1800
         Picture         =   "frmMain.frx":060D
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Example Notifications"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdExample1 
         Caption         =   "Alert"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdExample2 
         Caption         =   "Message"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdExample3 
         Caption         =   "Notification"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Timer tmrShowPopup 
      Interval        =   500
      Left            =   4200
      Top             =   2640
   End
   Begin VB.Label lblNotificationCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pending Notification Requests: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'Module     : frmMain
'Description:
'Version    : V2.00 1/11/2005 14:03
'Release    : VB6
'Copyright  :
'Author     : Chris.Nillissen
'----------------------------------------------------------------------------------
'V2.00    1/11/2005 Original version
'
'----------------------------------------------------------------------------------

Option Explicit


Public WithEvents m_NotificationWindow      As frmNotification
Attribute m_NotificationWindow.VB_VarHelpID = -1


Private Sub Command1_Click()
    Call mNotificationSystem.RequestUserNotification(txtKey.Text, txtTitle.Text, txtDescription.Text, (chkAllowSameKeys.Value = 0), (chkEnableClick.Value = 1), imgIcon(IIf(optIcon(0).Value, 0, 1)), txtSoundFileLocation.Text)
End Sub

Private Sub Form_Load()
    ' Create new instance of the Notification Requests Class.
    Set g_NotificationRequests = New Collection
    
    txtSoundFileLocation.Text = App.Path & "\Sounds\Message.wav"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' Destroy all used classes.
    Set g_NotificationRequests = Nothing
End Sub


Private Sub cmdExample1_Click()
Static m_AlertCount As Long
    If (m_AlertCount = 0) Then m_AlertCount = 10
    m_AlertCount = m_AlertCount - 1
    Call mNotificationSystem.RequestUserNotification("ALERT", "System Alert", "Warning! Your battery is at " & CStr(m_AlertCount) & " percent", False, True, imgIcon(0))
End Sub
Private Sub cmdExample2_Click()
Static m_MessageCount As Long

    m_MessageCount = m_MessageCount + 1
    Call mNotificationSystem.RequestUserNotification("MESSAGE", "New Message", "You have " & CStr(m_MessageCount) & " Unread Messages", False, True, imgIcon(1))
End Sub
Private Sub cmdExample3_Click()
Static m_ComputerCount As Long

    m_ComputerCount = m_ComputerCount + 1
    Call mNotificationSystem.RequestUserNotification("NOTIFY:" & "Host" & CStr(m_ComputerCount), "Notification", "The workstation 'Computer-" & CStr(m_ComputerCount) & "' as changed its resolution.", True, True)
End Sub



Private Sub m_NotificationWindow_Clicked(Key As String)
     ' *** Enter your own click handling code below. ***
     
     Select Case Key
     Case "ALERT":      MsgBox "You have clicked the Alerts Notification"
     Case "MESSAGE":    MsgBox "You have clicked the Messages Notification"
     Case Else:         MsgBox "You have clicked the Notification with Key: " & Key
     End Select
End Sub
Private Sub m_NotificationWindow_Finished()
    ' Set the Notification variable to nothing, to indicate that we have
    ' finished using it.
    Set m_NotificationWindow = Nothing
End Sub


Private Sub tmrShowPopup_Timer()
Dim lNotificationRequest      As cNotificationRequest

    ' Check if we have some requests and make sure we arent showing a notification already.
    If (g_NotificationRequests.Count > 0) And (m_NotificationWindow Is Nothing) Then  '(Not IsFormLoaded("frmUserPopup")) Then
        ' Get the first Notification Request from the Collection.
        Set lNotificationRequest = g_NotificationRequests.Item(1)
        
        ' Setup and Show the notification request.
        Set m_NotificationWindow = New frmNotification
        Call m_NotificationWindow.ShowNotification(lNotificationRequest)
        
        ' Remove the Request from the Collection.
        g_NotificationRequests.Remove 1
    End If
    
    ' Update Notification Count.
    lblNotificationCount.Caption = "Pending Notification Requests: " & CStr(g_NotificationRequests.Count)
        
    Set lNotificationRequest = Nothing
End Sub
