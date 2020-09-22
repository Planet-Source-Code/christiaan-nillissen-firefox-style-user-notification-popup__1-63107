VERSION 5.00
Begin VB.Form frmNotification 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   690
   ClientLeft      =   60
   ClientTop       =   75
   ClientWidth     =   3300
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
   MinButton       =   0   'False
   ScaleHeight     =   46
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrPopupController 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2880
      Top             =   45
   End
   Begin VB.Shape shpBorder 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   795
   End
   Begin VB.Image imgNotificationIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmNotification.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'Module     : frmNotification
'Description:
'Version    : V2.00 31/10/2005 10:17
'Release    : VB6
'Copyright  :
'Author     : Chris.Nillissen
'----------------------------------------------------------------------------------
'V2.00    31/10/2005 Original version
'
'----------------------------------------------------------------------------------

Option Explicit


Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long


Private Const SPI_GETWORKAREA    As Long = 48


Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type


Public Enum eStatusType
    StatusShow = 0
    StatusHide = 1
End Enum


Public Event Clicked(ByRef Key As String)
Public Event Finished()


Private m_Status                As eStatusType
Private m_FormOpenHeight        As Long
Private m_FormBottomPosition    As Long
Private m_FormRightPosition     As Long
Private m_OpenInterval          As Long


Private m_NotificationRequest   As cNotificationRequest


Private Sub Form_Load()
Dim lDesktopArea        As RECT
    
    ' Set default values.
    m_OpenInterval = 5000
    
    ' Set the Window as top most window.
    Call mNotificationSystem.SetWindowTopMost(Me.hwnd)
    
    ' Get desktop area not taken up by the taskbar.
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, lDesktopArea, 0&)
    m_FormOpenHeight = Me.Height
    m_FormBottomPosition = (lDesktopArea.Bottom * Screen.TwipsPerPixelY)
    m_FormRightPosition = (lDesktopArea.Right * Screen.TwipsPerPixelX)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Disable the Popup Controller Timer.
    tmrPopupController.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Finished
End Sub


Public Property Get NotificationRequest() As cNotificationRequest
    Set NotificationRequest = m_NotificationRequest
End Property
Public Property Let NotificationRequest(ByVal vNewValue As cNotificationRequest)
    m_NotificationRequest = vNewValue
End Property


Private Sub lblDescription_Click()
    If (m_NotificationRequest.EnableClickEvent) Then RaiseEvent Clicked(m_NotificationRequest.Key)
End Sub

Private Sub tmrPopupController_Timer()

    Select Case m_Status
    Case StatusShow
        
        Me.Move Me.Left, m_FormBottomPosition - Me.Height, Me.Width, Me.Height + 20
        If (Me.Height >= m_FormOpenHeight) Then
            Me.Height = m_FormOpenHeight
            m_Status = StatusHide
            tmrPopupController.Interval = m_OpenInterval
            Exit Sub
        End If
    
    Case StatusHide
        tmrPopupController.Interval = 15
        Me.Move Me.Left, m_FormBottomPosition - Me.Height, Me.Width, Me.Height - 20
        If (Me.Height < 20) Then Unload Me

    End Select

End Sub


Public Sub ShowNotification(ByVal NotificationRequest As cNotificationRequest)
    ' Store a copy of the Notification Request.
    Set m_NotificationRequest = NotificationRequest
    
    ' Setup the Window with the Notification Request settings.
    Call SetupNotification(NotificationRequest)

    ' Set starting position, size and show the window.
    Me.Move m_FormRightPosition - (Me.Width + 100), m_FormBottomPosition - 10, Me.Width, 10
    Me.Show: DoEvents
    
    ' Start showing the form starting at top of task bar.
    m_Status = StatusShow
    tmrPopupController.Enabled = True
    
    ' Play the associated wave file.
    Call mNotificationSystem.PlayWaveSoundFile(NotificationRequest.SoundFileLocation)
End Sub
Public Sub UpdateNotification(ByVal NotificationRequest As cNotificationRequest)
    ' Store a copy of the Notification Request.
    Set m_NotificationRequest = NotificationRequest
    
    ' Setup the Window with the Notification Request settings.
    Call SetupNotification(NotificationRequest)
    
    ' Start showing the form starting at top of task bar.
    m_Status = StatusShow
    tmrPopupController.Enabled = True
End Sub


Private Sub SetupNotification(ByRef NotificationRequest As cNotificationRequest)
    ' Setup the Forms Controls.
    lblTitle.Caption = NotificationRequest.Title
    lblDescription.Caption = NotificationRequest.Description
    If (Not NotificationRequest.Icon Is Nothing) Then Set imgNotificationIcon = NotificationRequest.Icon
    
    ' Size the window to fit the description.
    Me.Width = (lblDescription.Left + lblDescription.Width + 10) * Screen.TwipsPerPixelX
    Me.Left = m_FormRightPosition - (Me.Width + 100)
    
    ' Position any controls on the form.
    shpBorder.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
