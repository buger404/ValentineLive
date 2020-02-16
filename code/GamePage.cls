VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim RX As Single, RY As Single, Step As Long, LTime As Long
    Dim DX As Single, TDX As Single
    Dim People As Long, PDC As ScrollArea, lPeople As Long
    Dim pcTime As Long, pcMax As Long, pMax As Long
    Private Type CmtItem
        CmT As String
        time As Long
        y As Single
        X As Single
        Color As Long
        w As Long
    End Type
    Private Type Gun
        X As Single
        y As Single
        XBuff As Single
        YBuff As Single
        time As Long
        pic As String
        XR As Single
        YR As Single
    End Type
    Private Type Tip
        time As Long
        pic As String
    End Type
    Dim Crash As New GCrashBox, Crash2 As New GCrashBox
    Dim Guns() As Gun, MyGun() As Gun, GIndex As Integer
    Dim CmtItems() As CmtItem, Tips() As Tip
    Dim HP As Long, UHP As Long, ShakeTime As Long
    Dim lmX As Long, lmY As Long
    Dim LifeStep As Long, ATK As Long
    Dim QuietTime As Long, QMark As Boolean, QuietDelay As Long
    Dim PraiseTime As Long, GunT As Long, CmT As Long, lHP As Long, lUHP As Long
    Dim MGunT As Long, MGunTD As Long
    Dim LiveTime As Long, LivePic As String, LiveDelay As Long
'======================================
Public Sub NewTip(tipn As String)
    ReDim Preserve Tips(UBound(Tips) + 1)
    With Tips(UBound(Tips))
        .time = GetTickCount
        .pic = tipn
    End With
End Sub
Public Sub NewCmt()
    CmT = CmT + 1
    Randomize
    ReDim Preserve CmtItems(UBound(CmtItems) + 1)
    With CmtItems(UBound(CmtItems))
        .y = Rnd * (GH - 50)
        .time = GetTickCount
        Select Case Int(Rnd * 6)
            Case 0: .Color = argb(255, 128, 128, 128)
            Case 1: .Color = argb(255, 255, 0, 0)
            Case 2: .Color = argb(255, 0, 255, 0)
            Case 3: .Color = argb(255, 0, 0, 255)
            Case 4: .Color = argb(255, 255, 255, 0)
            Case 5: .Color = argb(255, 0, 255, 0)
        End Select
        Select Case Int(Rnd * 40)
            Case 0: .CmT = "赶上直播"
            Case 1: .CmT = "shit♂冰棍"
            Case 2: .CmT = "冰棍姐姐好漂亮"
            Case 3: .CmT = "爱了爱了"
            Case 4: .CmT = "挡住不让你们看我的冰棍"
            Case 5: .CmT = "我好喜欢冰棍啊"
            Case 6: .CmT = "冰棍的腿毛真靓"
            Case 7: .CmT = "冰棍好帅气"
            Case 8: .CmT = "冰棍520": People = People + 200: NewTip "pop.png"
            Case 9: .CmT = "我爱你！！！！！！！！！"
            Case 10: .CmT = "性感冰棍 在线发牌"
            Case 11: .CmT = "冰棍的声音好好听"
            Case 12: .CmT = "冰棍是我的 你们走开"
            Case 13: .CmT = "冰棍酱无敌"
            Case 14: .CmT = "奥利给"
            Case 15: .CmT = "最爱冰棍"
            Case 16: .CmT = "冰棍！！！嫁给我！"
            Case 17: .CmT = "我的冰棍不让你们看"
            Case 18: .CmT = "太帅了吧"
            Case 19: .CmT = "整个屏幕的帅！"
            Case 20: .CmT = "帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅帅"
            Case 21: .CmT = "棍棍！！！"
            Case 22: .CmT = "在床上和冰棍玩"
            Case 23: .CmT = "全网冰棍最帅"
            Case 24: .CmT = "性感冰棍 在线发牌"
            Case 25: .CmT = "冰棍冰棍冰棍"
            Case 26: .CmT = "这谁顶得住啊"
            Case 27: .CmT = "别说了 冰棍是我的"
            Case 28: .CmT = "前面那个说冰棍是自己的 你配不上"
            Case 29: .CmT = "冰棍！！！"
            Case 30: .CmT = "发现冰棍github了"
            Case 31: .CmT = "冰棍的QQ是1257472418"
            Case 32: .CmT = "加棍棍QQ了"
            Case 33: .CmT = "♂"
            Case 34: .CmT = "冰棍の腿毛"
            Case 35: .CmT = "拖控件大法！！！"
            Case 36: .CmT = "冰棍编程超厉害"
            Case 37: .CmT = "冰棍真是全能鬼才"
            Case 38: .CmT = "天才冰棍"
            Case 39: .CmT = "yes♂sir"
        End Select
        .w = EF.GetWidth(Page.GG, .CmT, 24, StringAlignmentNear, FontStyleBold)
        .X = Int(Rnd * GW)
    End With
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布
    
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    'Exit Sub
    
    Dim pro As Single
    If Step = 0 Then
        If GetTickCount - LTime >= 2000 Then
            ECore.IsLayered = False
            SetWindowLongA GHwnd, GWL_EXSTYLE, OldLong
            'ECore.FancyMode = True
            SE.Play "save.wav"
            LTime = GetTickCount: Step = 1
        End If
    End If
    
    Dim spro As Single
    If Step > 1 Then spro = 1
    
    If Step = 1 Then
        pro = (GetTickCount - LTime) / 1000
        If pro > 1 Then
            pro = 1: Step = 2
        End If
        spro = pro
        RY = GH / 2 + pro * GH * 0.25
    End If
    
    Dim Quiet As Boolean
    
    If Step > 1 Then
        Dim p As POINTAPI
        GetCursorPos p
        Quiet = (lmX = p.X And lmY = p.y)
        lmX = p.X: lmY = p.y
        RX = RX + (p.X - RX) / 7: RY = RY + ((p.y + 12) - RY) / 7
    End If
    
    Dim cs As Long
    For I = 1 To UBound(CmtItems)
redo:
        If I > UBound(CmtItems) Then Exit For
        With CmtItems(I)
            pro = (GetTickCount - .time) / 8000
            If pro > 1 Then pro = 1
            cs = GW - pro * (GW + .w + .X) + .X
            If cs < GW Then Page.Writes .CmT, cs, .y, 24, .Color, style:=FontStyleBold
        End With
        If pro = 1 Then CmtItems(I) = CmtItems(UBound(CmtItems)): ReDim Preserve CmtItems(UBound(CmtItems) - 1): GoTo redo
    Next
    
    Dim yy As Long, yy2 As Long
    
    If People > pcMax Then pcMax = (pcMax + People) / 2
    If GetTickCount - pcTime >= 3000 Then
        pMax = pcMax: pcTime = GetTickCount: pcMax = 1
        For I = 1 To Int(pMax / 1000)
            Call NewCmt
        Next
    End If
    
    If GunList(GIndex).time <> 0 Then
        If BGM.position >= GunList(GIndex).time Then
            Call NewGun(GunList(GIndex).Kind)
            GIndex = GIndex + 1
        End If
    End If
    
    Dim tHurt As Long
    tHurt = HP
    
    For I = 1 To UBound(Guns)
regun:
        If I > UBound(Guns) Then Exit For
        Select Case Guns(I).pic
            Case "b1.png"
                If (GetTickCount Mod 100 < 50) And (Guns(I).XR <> 0) Then
                    Page.Paint 1, Guns(I).XR, Guns(I).YR, 200, 200, argb(255, 255, 0, 0), style:=1, size:=3, pos:=posOnCenter
                End If
                If GetTickCount - Guns(I).time > 1200 Then
                    If Guns(I).time <> 0 Then Guns(I).time = 0: SE.Play "attack4.wav"
                    Guns(I).X = Guns(I).X + Guns(I).XBuff: Guns(I).y = Guns(I).y + Guns(I).YBuff
                    Page.DrawImage Guns(I).pic, Guns(I).X, Guns(I).y, alpha:=1, pos:=posOnCenter
                    If Crash.IsCrash Then HP = HP - 2: SE.Play "hurt.wav"
                End If
            Case "re.png"
                Guns(I).X = Guns(I).X + Guns(I).XBuff: Guns(I).y = Guns(I).y + Guns(I).YBuff
                Page.DrawImage Guns(I).pic, Guns(I).X, Guns(I).y, alpha:=1, pos:=posOnCenter
                If Crash.IsCrash Then
                    HP = HP + 20: SE.Play "recover.wav"
                    Guns(I).y = GH * 2
                End If
            Case Else
                If InStr(Guns(I).pic, "tm") > 0 Then
                    pro = (GetTickCount - Guns(I).time) / 500
                    If pro > 1 And pro < 2 Then SE.Play "source.wav": Guns(I).time = Guns(I).time - 500
                    If pro > 4 Then Guns(I).y = -1
                    If pro > 3 Then pro = 4 - pro
                    If pro > 1 Then pro = 1
                    If pro < 0 Then pro = 0
                    Page.DrawImage Guns(I).pic, Guns(I).X, Guns(I).y, alpha:=pro
                    If pro = 1 Then
                        If Crash.IsCrash Then HP = HP - 2: SE.Play "hurt.wav"
                    End If
                Else
                    Page.DrawImage Guns(I).pic, Guns(I).X, Guns(I).y, alpha:=1, pos:=posOnCenter
                    If Guns(I).pic = "a4.png" Then
                        If Crash.IsCrash = True And Quiet = True Then HP = HP - 1: SE.Play "hurt.wav"
                    ElseIf Guns(I).pic = "a5.png" Then
                        If Crash.IsCrash = True And Quiet = False Then HP = HP - 1: SE.Play "hurt.wav"
                    Else
                        If Crash.IsCrash Then HP = HP - 1: SE.Play "hurt.wav"
                    End If
                    Guns(I).X = Guns(I).X + Guns(I).XBuff: Guns(I).y = Guns(I).y + Guns(I).YBuff
                End If
        End Select
        If Guns(I).X < 0 Or Guns(I).y < 0 Or Guns(I).X > GW Or Guns(I).y > GH Then
            Guns(I) = Guns(UBound(Guns)): ReDim Preserve Guns(UBound(Guns) - 1): GoTo regun
        End If
    Next
    
    tHurt = tHurt - HP
    If tHurt >= 3 And tHurt <= 4 Then People = People + 100: NewTip "violence.png"
    If tHurt >= 4 And tHurt <= 5 Then People = People + 200: NewTip "blood.png"
    If tHurt >= 8 Then People = People - 1000: NewTip "dying.png"
    If tHurt >= 3 And HP < 100 Then People = People + 300: NewTip "scream.png"
    If tHurt >= 3 And HP < 50 Then People = People + 400: NewTip "struggle.png"
    
    For I = 1 To UBound(MyGun)
regun2:
        If I > UBound(MyGun) Then Exit For
        Page.DrawImage MyGun(I).pic, MyGun(I).X, MyGun(I).y, alpha:=1, pos:=posOnCenter
        If Crash2.IsCrash Then
            MGunTD = MGunTD + 1
            UHP = UHP - 27 * Cubic((People / 23333), 0, 1, 1, 1) * Int(Rnd * 0.3 + 1) + 3: MyGun(I).y = GH * 2
            ShakeTime = GetTickCount
        End If
        MyGun(I).X = MyGun(I).X + MyGun(I).XBuff: MyGun(I).y = MyGun(I).y + MyGun(I).YBuff
        If MyGun(I).X < 0 Or MyGun(I).y < 0 Or MyGun(I).X > GW Or MyGun(I).y > GH Then
            MyGun(I) = MyGun(UBound(MyGun)): ReDim Preserve MyGun(UBound(MyGun) - 1): GoTo regun2
        End If
    Next
    
    If CheckMouse(0, 0, GW, GH) = mMouseUp Then
        ReDim Preserve MyGun(UBound(MyGun) + 1)
        With MyGun(UBound(MyGun))
            .X = RX
            .y = RY
            .YBuff = (GH / 2 - .y) / 7
            .pic = "a" & IIf(.YBuff < 0, 6, 3) & ".png"
        End With
        MGunT = MGunT + 1
        SE.Play "gun.wav"
        If Mouse.button = 2 Then Call NewTip("dying.png")
    End If
    
    If HP > 200 Then HP = 200
    If HP < 0 Then
        HP = 0
        BGM.Stops
        Page.Paint 0, 0, 0, GW, GH, argb(140, 255, 0, 0)
        Page.Paint 0, GW / 2 - 250, GH - 30 - 40, 500, 30, argb(255, 128, 128, 128)
        Page.Paint 0, GW / 2 - 250, GH - 30 - 40, 500 * (HP / 200), 30, argb(255, 0, 255, 0)
        Page.Writes HP & "/200", GW / 2, GH - 30 - 40, 18, argb(255, 255, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
        BitBlt GameWindow.hdc, 0, 0, GW, GH, Page.CDC, 0, 0, vbSrcCopy
        SE.Play "heartbreak.wav"
        For I = 1 To 150
            Sleep 10: DoEvents
        Next
        SE.Play "heartbreak2.wav"
        For I = 1 To 200
            Sleep 10: DoEvents
        Next
        BGM.Create App.path & "\music\GameOver.mp3"
        BGM.Play
        ECore.SimpleMsg "你失败了", "直播结束", StrArray("好吧")
        Unload GameWindow
        End
    End If
    
    If People < 0 Then People = 0
    
    Dim dPeople As Long
    dPeople = People * (1 + Rnd * 0.4 - 0.2)

    StartScrollArea Page, PDC
        BitBlt PDC.DC, -5, 0, 300, 140, PDC.DC, 0, 0, vbSrcCopy
        yy = 140 - dPeople / pMax * 70
        yy2 = lPeople
        lPeople = yy
        Page.Paint 0, 295, 0, 6, 141, argb(255, 64, 64, 64)
        Page.PaintLine 294, yy2, 300, yy, argb(255, 0, 255, 0), 2
    EndScrollArea Page, 40, 40, 0, 0, 300, 140, alpha:=spro
    
    Dim TMode As Integer, y As Long, dMark As Boolean
    y = 290
    'HP = 200
    
    For I = 1 To UBound(Tips)
retip:
        dMark = False
        If I > UBound(Tips) Then Exit For
        pro = (GetTickCount - Tips(I).time) / 1000
        If pro < 1 Then pro = Cubic(pro, 0, 1, 1, 1)
        If pro >= 1 And pro <= 4 Then pro = 1
        If pro > 5 Then dMark = True: pro = 0
        If pro >= 4 And pro <= 5 Then pro = 1 - Cubic(pro - 4, 0, 1, 1, 1)
        If TMode = 1 Then
            Page.DrawImage Tips(I).pic, GW - 40 * pro, y, alpha:=pro, pos:=posOnRight
        Else
            Page.DrawImage Tips(I).pic, 40 * pro, y, alpha:=pro
        End If
        y = y + 60 * pro
        If y >= GH Then
            If TMode = 1 Then Exit For
            y = 40 + 168 + 40: TMode = 1
        End If
        If dMark Then
            For S = I + 1 To UBound(Tips)
                Tips(S - 1) = Tips(S)
            Next
            ReDim Preserve Tips(UBound(Tips) - 1)
            GoTo retip
        End If
    Next
    
    If GetTickCount - PraiseTime >= 5000 Then
        If HP < 50 Then People = People + 400: NewTip "struggle.png"
        If CmT > 50 Then People = People + 200: NewTip "pop.png"
        If GunT > 50 Then People = People + 200: NewTip "fancy.png"
        If lHP - HP >= 1000 Then People = People + 1500: NewTip "effort.png"
        If lHP - HP <= -400 Then People = People + 500: NewTip "relife.png"
        If lUHP - lUHP >= 100 Then People = People + 200: NewTip "courage.png"
        If lUHP - lUHP >= 200 Then People = People + 600: NewTip "anger.png"
        If MGunT > 0 Then
            If MGunTD / MGunT >= 0.9 Then People = People + 400: NewTip "gun.png"
        End If
        PraiseTime = GetTickCount: GunT = 0: CmT = 0: lHP = HP: lUHP = UHP
        MGunT = 0: MGunTD = 0
    End If
    
    Page.Paint 0, 40, 180, 301, 100, argb(70 * spro + 1, 25, 255, 255)
    Page.Writes "收视率 " & Int(pMax * spro), 50, 190, 24, argb(254 * spro + 1, 255, 255, 255), style:=FontStyleBold, BorderColor:=argb(254 * spro + 1, 0, 0, 0), BorderSize:=1
    Page.Writes "冰棍酱の直播间", 50, 230, 24, argb(254 * spro + 1, 255, 255, 255), style:=FontStyleBold, BorderColor:=argb(254 * spro + 1, 0, 0, 0), BorderSize:=1
    
    'Page.Paint 0, 40, 40, 301, 140 + 100, argb(128 * spro + 1, 0, 0, 0), style:=1, size:=2
    
    Page.DrawImage "h1.png", RX, RY, alpha:=1, pos:=posOnCenter
    Crash.Reset
    DX = DX + (TDX - DX) / 10
    If GetTickCount - ShakeTime <= 1000 Then
        Page.DrawImage "step" & LifeStep & ".png", DX + Int(Rnd * 10 - 5), GH / 2 + Int(Rnd * 10 - 5), alpha:=1, pos:=posOnCenter
    Else
        Page.DrawImage "step" & LifeStep & ".png", DX, GH / 2, alpha:=1, pos:=posOnCenter
    End If
    Crash2.Reset
    
    Dim LUMark As Boolean
    If LifeStep < 2 And UHP < 200 Then UHP = 200
    If UHP < 700 And LifeStep = 0 And People >= 4444 Then LUMark = True
    If UHP < 400 And LifeStep = 1 And People >= 15555 Then LUMark = True
    If UHP < 0 And LifeStep = 2 Then
        BGM.Stops
        Sleep 3000
        ECore.NewTransform transFallDark, 3000, "EndPage"
    End If
    People = People + 1
    
    If LUMark Then
        People = People + 2000: NewTip "boom.png"
        SE.Play "Thunder9.mp3"
        ECore.NewTransform transHighLight, 500
        UHP = 1000: LifeStep = LifeStep + 1
    End If
    
    If LifeStep = 2 Then
        pro = 1 + Cubic((1 - (UHP / 1000)), 0, 1, 1, 1) * 0.5
        If BGM.rate <> pro Then BGM.rate = pro
    End If
    
    If GetTickCount - LiveTime >= LiveDelay Then
        Randomize
        LiveTime = GetTickCount
        Select Case LifeStep
            Case 0
                If Int(Rnd * 4) = 1 Then
                    LivePic = "girl" & Int(Rnd * 8) + 1 & ".jpg": LiveDelay = 100
                Else
                    LivePic = "boy" & Int(Rnd * 5) & ".png": LiveDelay = 1000
                End If
            Case 1
                If Int(Rnd * 2) = 1 Then
                    LivePic = "girl" & Int(Rnd * 8) + 1 & ".jpg": LiveDelay = 500
                Else
                    LivePic = "boy" & Int(Rnd * 5) & ".png": LiveDelay = 500
                End If
            Case 2: LivePic = "girl" & Int(Rnd * 8) + 1 & ".jpg": LiveDelay = 1000
        End Select
    End If
    
    Page.DrawImage LivePic, GW - 40 - 128, 40, alpha:=spro
    Page.Writes "冰棍酱の摄像头", GW - 40, 180, 24, argb(254 * spro + 1, 255, 255, 255), style:=FontStyleBold, BorderColor:=argb(254 * spro + 1, 0, 0, 0), BorderSize:=1, align:=StringAlignmentFar
    
    Page.Paint 0, GW / 2 - 250, GH - 30 - 40, 500, 30, argb(255, 128, 128, 128)
    Page.Paint 0, GW / 2 - 250, GH - 30 - 40, 500 * (HP / 200), 30, argb(255, 0, 255, 0)
    Page.Writes HP & "/200", GW / 2, GH - 30 - 40, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
    
    Page.Paint 0, GW / 2 - 250, 40, 500, 30, argb(255, 128, 128, 128)
    Page.Paint 0, GW / 2 - 250, 40, 500 * (UHP / 1000), 30, argb(255, 255, 0, 0)
    If LifeStep < 2 And UHP <= 200 Then
        Page.Writes "无敌状态", GW / 2, 40 + 3, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
    End If
    
    If BGM.IsPlaying = False Then
        BGM.Play: GIndex = 1
    End If
End Sub
Public Sub NewGun(index As Integer)
    GunT = GunT + 1

    If index <> 5 Then
        If LifeStep = 0 Then index = 1
        If LifeStep = 1 And index <= 1 Then index = 4
        If LifeStep = 1 And index = 2 Then index = 5
    End If
    
    Select Case index
        Case 0
            SE.Play "attack3.wav"
            For I = 0 To 7
                ReDim Preserve Guns(UBound(Guns) + 1)
                With Guns(UBound(Guns))
                    .pic = "a1.png"
                    .time = GetTickCount
                    .X = DX
                    .y = GH / 2
                    '0=+0,1=0+,2=++,3=+-,4=-+,5=-0,6=0-,7=--
                    .XBuff = 5 * IIf(I = 0 Or I = 2 Or I = 3, 1, -1) * IIf(I = 1 Or I = 6, 0, 1)
                    .YBuff = 5 * IIf(I = 1 Or I = 2 Or I = 4, 1, -1) * IIf(I = 0 Or I = 5, 0, 1)
                End With
            Next
        Case 1
            SE.Play "attack4.wav"
            For I = 0 To 3
                ReDim Preserve Guns(UBound(Guns) + 1)
                With Guns(UBound(Guns))
                    .pic = "a0.png"
                    .time = GetTickCount
                    .X = DX
                    .y = GH / 2
                    .XBuff = (RX - .X) / 40 * IIf(I = 0 Or I = 2, 1, -1)
                    .YBuff = (RY - .y) / 40 * IIf(I = 1 Or I = 2, 1, -1)
                    If LifeStep > 0 Then
                        If Abs(.XBuff) < 2 Then .XBuff = 2 * IIf(.XBuff < 0, -1, 1)
                        If Abs(.YBuff) < 2 Then .YBuff = 2 * IIf(.YBuff < 0, -1, 1)
                    End If
                End With
            Next
        Case 2
            SE.Play "attack5.wav"
            For I = 0 To 3
                ReDim Preserve Guns(UBound(Guns) + 1)
                With Guns(UBound(Guns))
                    .pic = "b1.png"
                    .time = GetTickCount
                    .X = RX + 200 * IIf(I = 1 Or I = 3, 1, -1)
                    .y = RY + 200 * IIf(I = 0 Or I = 3, 1, -1)
                    .XBuff = 20 * IIf(I = 0 Or I = 2, 1, -1)
                    .YBuff = 20 * IIf(I = 1 Or I = 2, 1, -1)
                End With
                Guns(UBound(Guns)).XR = RX: Guns(UBound(Guns)).YR = RY
            Next
        Case 3
            Randomize
            Dim basex As Long
            basex = RX - 50
            For I = 0 To 3
                ReDim Preserve Guns(UBound(Guns) + 1)
                With Guns(UBound(Guns))
                    .pic = "tm" & Int(Rnd * 6) & ".png"
                    .time = GetTickCount + I * 100
                    .X = basex + Int(Rnd * 100 - 50)
                    .y = 0
                    .XBuff = 0
                    .YBuff = 0
                End With
                Guns(UBound(Guns)).XR = RX: Guns(UBound(Guns)).YR = RY
            Next
        Case 4
            Randomize
            Dim Mark As Long
            Mark = Int(Rnd * 2)
            SE.Play "attack1.wav"
            If GetTickCount - QuietTime >= QuietDelay Then
                QuietDelay = Int(Rnd * 2000 + 2000)
                QMark = Not QMark: QuietTime = GetTickCount
            End If
            For I = 0 To 36
                ReDim Preserve Guns(UBound(Guns) + 1)
                With Guns(UBound(Guns))
                    .pic = "a" & IIf(QMark, 4, 5) & ".png"
                    .time = GetTickCount
                    .X = DX
                    .y = GH / 2
                    .XBuff = Cos(I / 36 * 360) * 10
                    .YBuff = Sin(I / 36 * 360) * 10
                End With
            Next
        Case 5
            TDX = Rnd * GW * 0.7 + GW * 0.15
    End Select
    
    If Int(Rnd * 30) > 30 - (1 - (HP / 200)) * 4 * IIf(LifeStep = 2, 2, 1) Then
        ReDim Preserve Guns(UBound(Guns) + 1)
        With Guns(UBound(Guns))
            .pic = "re.png"
            .time = GetTickCount
            .X = Rnd * GW
            .y = 0
            .YBuff = 5
        End With
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    ReDim Guns(0)
    ReDim MyGun(0)
    ReDim Tips(0)
    LifeStep = 0
    DX = GW / 2: TDX = GW / 2
    RX = GW / 2: RY = GH / 2: LTime = GetTickCount
    pcMax = 1: pMax = 1: GIndex = 1
    People = 1000
    BGM.Create App.path & "\music\Icelolly's special live.mp3"
    BGM.Volume = 0.5
    BGM.Play
    HP = 200: UHP = 1000
    
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets", 0.35
    Page.Res.NewImages App.path & "\assets\tm", GH / 700
    Page.Res.NewImages App.path & "\assets\tip", 0.7
    Page.Res.NewImages App.path & "\assets\live", 128, 128
    PDC = CreateScrollArea(300, 140)
    '创建页面
    ECore.Add Page, "GamePage"
    ReDim CmtItems(0)
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
