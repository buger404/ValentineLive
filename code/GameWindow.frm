VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Icelolly Special Live"
   ClientHeight    =   6672
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   ҳ�������
    Dim EC As GMan, CloseMark As Boolean
'==================================================
'   �ڴ˴��������ҳ����ģ������
'   Happy Valentine's Day
    Dim StartupPage As StartupPage
    Dim GamePage As GamePage
    Dim EndPage As EndPage
'==================================================

Private Sub DrawTimer_Timer()
    '����
    EC.Display
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Exit Sub
    If KeyCode = vbKeyF2 Then
        Dim sep() As GunKey
        ReDim sep(0)
        
        For I = 1001 To UBound(GunList)
            ReDim Preserve sep(UBound(sep) + 1)
            sep(UBound(sep)) = GunList(I)
        Next
        
        Open App.path & "\note\notes.love" For Binary As #1
        Get #1, , sep
        Close #1
    End If
    
    ReDim Preserve GunList(UBound(GunList) + 1)
    With GunList(UBound(GunList))
        .time = BGM.position
        Select Case KeyCode
            Case vbKeyA: .Kind = 0
            Case vbKeyS: .Kind = 1
            Case vbKeyD: .Kind = 2
            Case vbKeyF: .Kind = 3
            Case vbKeyJ: .Kind = 4
            Case vbKeyK: .Kind = 5
        End Select
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And SysPage.MsgButton <> -1 Then
        If ECore.SimpleMsg("���Ҫ�뿪ֱ������", "o(�i�n�i)o", StrArray("�ϻ�", "����")) = 1 Then Exit Sub
        Unload Me
        End
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then MsgBox "�����ѿ���": End

    ReDim GunList(1000)
    Open App.path & "\note\notes.love" For Binary As #1
    Get #1, , GunList
    Close #1
    'ReDim GunList(0)
    
    '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�СӴ~��
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX + 2, Screen.Height / Screen.TwipsPerPixelY + 2, False
    '��������
    MakeFont "΢���ź�"
    '����ҳ�������
    Set EC = New GMan
    OldLong = GetWindowLongA(GHwnd, GWL_EXSTYLE)
    EC.Layered False
    'EF.RenderMode = 0
    'EC.FancyMode = True
    
    '�����浵����ѡ��
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б�
    Set SE = New GMusicList
    SE.Create App.path & "\sound"

    '��ʼ��ʾ
    'Me.Show
    'DrawTimer.Enabled = True
    
    '�ڴ˴���ʼ�����ҳ��
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set StartupPage = New StartupPage
        Set GamePage = New GamePage
        Set EndPage = New EndPage
    '=============================================

    '���ûҳ��
    EC.ActivePage = "StartupPage"
    
    'SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Me.Show

    Do
        EC.Display: DoEvents
    Loop Until CloseMark
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseMark = True
    '��ֹ����
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    EndEmerald
    End
End Sub
