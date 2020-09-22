Attribute VB_Name = "mNotificationSystem"
'----------------------------------------------------------------------------------
'Module     : mNotificationSystem
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


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Private Const SWP_NOMOVE         As Long = 2
Private Const SWP_NOSIZE         As Long = 1
Private Const FLAGS              As Long = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST       As Long = -1
Private Const HWND_NOTOPMOST     As Long = -2

Private Const SND_SYNC          As Long = &H0
Private Const SND_ASYNC         As Long = &H1
Private Const SND_NODEFAULT     As Long = &H2
Private Const SND_LOOP          As Long = &H8
Private Const SND_NOSTOP        As Long = &H10


Public g_NotificationRequests   As Collection


Public Sub RequestUserNotification(ByRef Key As String, ByRef Title As String, ByRef Description As String, ByRef AllowSameType As Boolean, ByRef EnableClickEvent As Boolean, Optional ByVal Icon As Variant = Nothing, Optional ByVal SoundFileLocation As String = vbNullString)
Dim lNotificationRequest        As cNotificationRequest
Dim lItemKey                    As String
Dim lAddToCollectionRequired    As Boolean
On Error Resume Next

    Set lNotificationRequest = New cNotificationRequest
    With lNotificationRequest
        .Key = Key
        .Title = Title
        .Description = Description
        .Icon = Icon
        .EnableClickEvent = EnableClickEvent
        .SoundFileLocation = SoundFileLocation
    End With
    
    lAddToCollectionRequired = True
    
    If (Not frmMain.m_NotificationWindow Is Nothing) And (Not AllowSameType) Then
        If (frmMain.m_NotificationWindow.NotificationRequest.Title = Title) Then
            Call frmMain.m_NotificationWindow.UpdateNotification(lNotificationRequest)
            lAddToCollectionRequired = False
        End If
    End If
        
    If (lAddToCollectionRequired) Then
        ' Build the Request Key.
        lItemKey = IIf(Not AllowSameType, Title, Description)
        ' Remove Duplicates and Add the UserPopup Request to the Collection.
        Call RemoveDuplicateRequest(lItemKey)
        g_NotificationRequests.Add lNotificationRequest, lItemKey
    End If
End Sub

Private Sub RemoveDuplicateRequest(ByRef Key As String)
On Error Resume Next
    ' Request a remove of an item with the same key.
    g_NotificationRequests.Remove Key
End Sub


Public Sub PlayWaveSoundFile(ByRef SoundFile As String)
On Error Resume Next
    ' Check if a sound file has been past.
    If (LenB(SoundFile) > 0) Then
        ' Play the selected wave file.
        Call sndPlaySound(SoundFile, SND_ASYNC Or SND_NODEFAULT)
    End If
End Sub

Public Sub SetWindowTopMost(ByRef Handle As Long)
On Error Resume Next
    ' Set the window to be top most.
    Call SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
