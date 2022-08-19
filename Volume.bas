Option Explicit
Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
   ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
   
    Const VK_VOLUME_MUTE = &HAD 'Volume Mute key
    Const VK_VOLUME_DOWN = &HAE 'Volume Down key
    Const VK_VOLUME_UP = &HAF 'Volume Up key
   

Sub VolUp()
'-- Turn volumn up --
   keybd_event VK_VOLUME_UP, 0, 1, 0
   keybd_event VK_VOLUME_UP, 0, 3, 0
End Sub

Sub VolMute()
'-- Toggle mute on / off --
   keybd_event VK_VOLUME_MUTE, 0, 1, 0
End Sub

