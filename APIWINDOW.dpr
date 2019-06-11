//前些天分析了启明公司的口语光盘，同时也用Delphi编写了个补丁程序。以往的补丁程序原理是利用了资源释放形式进行，因此，最终生成的补丁程序会奇大无比（＝自身体积加上目标破解文件的体积），就拿启明光盘的破解补丁说吧，生成出来并压缩后仍然有800K左右。但考虑到补丁原理只是修改了一个偏移量的值，如此大动干戈地把破解后的文件制作成资源文件来进行覆盖，实在是下策，有没有更好的办法呢。肯定有！！！于是就有了自己的偏移量补丁程序模块啦。毕竟作为注册机或补丁程序的话，体积越小自然越好，因此，考虑用纯API编写，这样生成的程序可以很小很小，以下就是源代码，生成的程序仅有54K，压缩后只有40多K，太可爱啦！！但界面一样。（如果不使用GUI界面而只使用关键破解代码的话，生成出来的文件只有7K.呵呵。这才是真正意义上的补丁哦。
//******************** 以下是需要修改的常量或变量　　　 ****************
//    （*主对话框界面*）
program apiwindow;
uses Windows,ShellAPI,Messages,System;
const
szMainCaption = '海浪轻风偏移量补丁程序V1.0';
szName = '启明英语口语考试系统'; //目标破解文件名
szBy = 'futuring@126.com'; //软件开发者
szCracker = '海浪轻风'; //解密人
szCrackeremail = 'mailto:futuring@126.com'; //解密人email
szLink = 'http://hi.baidu.com/beyond0769'; //主页地址
szTime = '2008-03-18'; //破解时间
szCopyright = '本程序只作研究学习用,禁止非法用途';
szKeyGenName = '海浪轻风偏移量补丁程序';
szScroll = '海浪轻风偏移量补丁程序' + #$0A + #$0A
          + '本模块用纯API函数编写，界面构建'+ #$0A + #$0A
          + '利用了资源文件形式，大大节省了'+ #$0A + #$0A
          + '生成后的文件体积，压缩后只有'+ #$0A + #$0A
          + '41k左右。同时加入了uFMOD单元' + #$0A + #$0A
          + '实现背景音乐循环播放。 ' + #$0A + #$0A + #$0A
          + '特别鸣谢：' + #$0A + #$0A
          + 'you_know[FCG]'+ #$0A+ #$0A
          + '看雪学院的Delphi牛人们' + #$0A+ #$0A
          + '以及uFMOD的程序员'+ #$0A+ #$0A
          + '感谢一直在背后默默支持我的琴儿' + #$0A + #$0A
          + '补丁使用方法：'+ #$0A + #$0A
          + '把补丁复制到目标文件相同目录下'+ #$0A + #$0A
          + '运行，点击“应用补丁”即可。 '+ #$0A + #$0A
          + '本程序只作研究学习用' + #$0A+#$0A
          + '禁止一切非法商业用途' + #$0A+#$0A
          + '程序破解 by 海浪轻风(黄仁来)' + #$0A+#$0A
          + '程序编程 by 海浪轻风(黄仁来)' + #$0A+#$0A
          + 'futuring@126.com' + #$0A + #$0A
          + 'http://beyond-0769.blog.163.com/' + #$0A + #$0A
          + '2008-03-18';
var
    FileName: PChar = 'SpokenEngMain.exe'; //破解目标文件完整名称
    IntFileSize: Cardinal = 2731008; //破解目标文件的大小字节
    RBuffer: array[0..1] of Byte = ($75, $0C); //目标破解文件原有的偏移量
    WBuffer: array[0..1] of Byte = ($EB, $0C); //修改后的偏移量
    OffsetPos: TOVERLAPPED = (Internal: 0; InternalHigh: 0; Offset: $000EE089; OffsetHigh: 0; hEvent: 0);
    CommandLine: PChar; // ↑↑以上是物理地址，通过W32DSM可以查到
    hwndname: HWND;
    hFile: THANDLE;
    Numb: DWORD;
    Buffer: array[0..1] of Byte;
    nType: DWORD;
    pMsg: string;
////////////////////
//全局变量数据声明 ******************** 以下常量或变量基本上无需要修改 ******************8
////////////////////
    h_Icon: HICON; //实例句柄
    h_Inst: HMODULE;
    h_Cur: hCursor;
    h_Brush: HBRUSH;
    sRectL, sRectM, sRectA: TRect;
 //RegName,RegCode:Array[1..256] of Byte; //原版数据；下句是修改后的
    RegName, RegCode: string;
    h_ScrollParent, h_Scroll: HWND;
    ScrollWidth, ScrollHeight: Word;
    xCount: Integer;
    LineCount: Word;
const
////////////////
//资源常量定义//
////////////////
    MAINICON = 100;
    LINKCURSOR = 112;

    IDD_LICENSEDLG = 1000;
    IDD_MAINDLG = 2000;
    IDD_ABOUTDLG = 3000;

    LICENSE_YES = 1001;
    LICENSE_NO = 1002;
    LICENSE_CLOSE = 1004;

    MAIN_CALC = 2001; //计算按钮控件ID（本实例已经去掉）
    MAIN_EXIT = 2002; //退出按钮控件ID
    MAIN_ABOUT = 2003; //关于按钮控件ID
    MAIN_CLOSE = 2004; //关闭按钮控件ID

    ABOUT_OK = 3001;
    ABOUT_CLOSE = 3002;

{$R dialog.res}
//++++++++++++++++++关键的爆破过程++++++++++++++++++++
procedure PatchFile;
var
    Res:boolean;
begin
    Setfileattributes(FileName, FILE_ATTRIBUTE_NORMAL + FILE_ATTRIBUTE_ARCHIVE); //设置文件的属性为正常
    hFile := CreateFile(FileName, GENERIC_READ or GENERIC_WRITE, FILE_SHARE_READ or
    FILE_SHARE_WRITE, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
    try
        if hFile <> INVALID_HANDLE_VALUE then
        begin
            if GetFileSize(hFile,nil)<>IntFileSize then
            begin
                nType := MB_OK or MB_ICONINFORMATION;
                pMsg := '文件校验出错：'+#13+'文件大小不一致，目标文件(英语口语主程序:' + FileName + ')可能已经被修改，拒绝打补丁！';
                exit;
            end;
            ReadFile(hFile, Buffer, 2, Numb, @OffsetPos); //如果读出的偏移数据是已经打上补丁的数据，则提示已经打上补丁。
            if Word(Buffer[0]) = Word(WBuffer[0]) then begin
                nType := MB_OK or MB_ICONINFORMATION;
                pMsg := '错误提示：'+#13+'目标程序(' + FileName + ')已经打过补丁了，还想怎样？ ^_^';
                Res:=True;
                exit;
                end;
            if Word(Buffer[0]) = Word(RBuffer[0]) then {// 读取偏移是否正确；} begin
                CopyFile(FileName, PChar('备份' + FileName), False); //备份破解目标文件；
                if WriteFile(hFile, WBuffer, 2, Numb, @OffsetPos) then
                begin
                    nType := MB_OK or MB_ICONINFORMATION;
                    pMsg := '成功打上补丁! Enjoy it^_^' + #13 +#13+
                            'Cracked & Coded by 海浪轻风(黄仁来)'; ; //成功补丁；
                    Res:=True;
                end else
                begin // 写入出错。
                    nType := MB_OK or MB_ICONERROR;
                    pMsg := '错误提示：'+#13+'：请检查目标文件(' + FileName + ')是否已经被打开，请先将其关闭再尝试。';
                end;
            end else {//偏移不正确则提示出错；}
            begin
                nType := MB_OK or MB_ICONWARNING;
                pMsg := '错误提示：'+#13+'目标文件(' + FileName + ')偏移量不符，拒绝打补丁！'+#13+'更多问题请到个人主页反映。';
                ShellExecute(0, nil, szLink, nil, nil, 0);
            end;
        end else {//文件被打开或文件找不到}
        begin
            nType := MB_OK or MB_ICONERROR;
            pMsg := '补丁出错 :-( 可能原因如下：' + #13 +
                    '1.目标文件找不到:(英语口语主程序:' + FileName + ')，请把本程序复制到同一个目录下再运行。' + #13 +
                    '2.目标文件已经被打开，请先将其关闭再尝试。';
        end;
    finally
        CloseHandle(hFile);
        MessageBox(0, PChar(pMsg), PChar('海浪轻风温馨提示：'), nType);
        if Res then ShellExecute(0, nil, szLink, nil, nil, 0);
    end;
end;


//+++++++++++++++++以下为构造程序主窗口所需的函数过程++++++++++++++

Function CountCRLF(Str:String):Word;
Var
    Count:Word;
    i:Word;
Begin
    Count:=1;
    For i:=1 to Length(Str) Do
    Begin
        If Str[i]=#$0A Then Inc(Count);
    End;
    CountCRLF:=Count;
End;

Procedure DialogInit(hDlg:HWND);
Var
    RECT:TRect;
    X,Y:Word;
    i:Word;
    Rgn:HRGN;
Begin
    ShowWindow(hDlg,SW_HIDE);
    GetWindowRect(hDlg,RECT);
    X:=(RECT.Right-RECT.Left) DIV 2;
    Y:=(RECT.Bottom-RECT.Top) DIV 2;

    For i:=1 to RECT.Right DIV 2 do
    Begin
        Rgn:=CreateRectRgn(X-i,Y-i,X+i,Y+i);
        SetWindowRgn(hDlg,Rgn,True);
        ShowWindow(hDlg,SW_SHOW);
        DeleteObject(Rgn);
    End;
End;

Procedure ItemDraw(lpDIS:PDrawItemStruct);
Var
    Str:Array[1..11] of Byte;
Begin
    FillRect(lpDIS^.hDC,lpDIS^.rcItem,h_Brush);
    SetTextColor(lpDIS^.hDC,$A0A0A0);
    //SetTextColor(lpDIS^.hDC,$03FF09); //绿色
    SetBkMode(lpDIS^.hDC,TRANSPARENT);
    DrawEdge(lpDIS^.hDC,lpDIS^.rcItem,BDR_RAISEDOUTER,BF_RECT);
    GetWindowText(lpDIS^.hwndItem,@Str,10);
    DrawText(lpDIS^.hDC,@Str,-1,lpDIS^.rcItem,DT_SINGLELINE OR DT_VCENTER OR DT_CENTER);

    If lpDIS^.itemState MOD 2=1 Then
    Begin
        SetTextColor(lpDIS^.hDC,$03FF09); //按下 绿色
        DrawText(lpDIS^.hDC,@Str,-1,lpDIS^.rcItem,DT_SINGLELINE OR DT_VCENTER OR DT_CENTER);
        DrawEdge(lpDIS^.hDC,lpDIS^.rcItem,BDR_SUNKENOUTER,BF_RECT);
    End;
End;
Procedure Paint(dc:HDC;Icon:hIcon;Caption:PChar;BeginRGB:DWORD;EndRGB:DWORD;Rect:TRect);
Var
    LogBrush:TLogBrush;
    LRect,rc:TRect;
    LhBR:HBRUSH;
    R,G,B,Ri,Gi,Bi,Rt,Gt,Bt:Word;
    LFont:HFONT;
    Width:Word;
    i:Word;
Begin
    LRect:=Rect;

    B:=EndRGB AND $FF;
    G:=(EndRGB SHR 8) AND $FF;
    R:=(EndRGB SHR 16) AND $FF;
    Bi:=BeginRGB AND $FF;
    Gi:=(BeginRGB SHR 8) AND $FF;
    Ri:=(BeginRGB SHR 16) AND $FF;
    Width:=(LRect.right-LRect.left) DIV 2;
    rc.top:=0;
    rc.bottom:=LRect.bottom-LRect.top;
    LogBrush.lbStyle:=0;
    LogBrush.lbHatch:=0;
    If Width>=0 Then
    Begin
        For i:=0 To (Width AND $FF) Do
        Begin
            rc.left:=1;
            rc.right:=i+1;
            Bt:=MulDiv(i,Bi,Width)+B;
            Gt:=MulDiv(i,Gi,Width)+G;
            Rt:=MulDiv(i,Ri,Width)+R;
            If Bt>$FF Then Bt:=$FF;
            If Gt>$FF Then Gt:=$FF;
            If Rt>$FF Then Rt:=$FF;
            LogBrush.lbColor:=(Rt SHL 16) OR (Gt SHL 8) OR (Bt);
            LhBR:=CreateBrushIndirect(LogBrush);
            FillRect(dc,rc,LhBR);
            rc.left:=Rect.right-Rect.left-i;
            rc.right:=Rect.right-Rect.left-i-1;
            FillRect(dc,rc,LhBR);
            DeleteObject(LhBR);
        End;
    End;
    LogBrush.lbColor:=$9E6A54; //标题栏颜色
    LhBR:=CreateBrushIndirect(LogBrush);
    FrameRect(dc,Rect,LhBR);
    DeleteObject(LhBR);
    SetTextColor(dc,$05E8FC);
    SetBkMode(dc,TRANSPARENT);
    Rect.left:=2;
    Rect.top:=2;
    Dec(Rect.bottom,2);
    LFont:=CreateFont(-$C,0,0,0,$2BC,0,0,0,1,0,0,0,0,'宋体');
    SelectObject(dc,LFont);
    If Icon<>0 Then
    Begin
        DrawIconEx(dc,2,2,Icon,$10,$10,0,0,3);
        Rect.left:=$14;
    End;
    DrawText(dc,Caption,-1,Rect,$24);
    DeleteObject(LFont);
End;

//////////////////////////////////////////////////////////////////

//滚动文本函数

Function ScrollProc(hDlg:HWND;Msg,wParam,lParam:DWORD):LRESULT;stdcall;
Var
    sRect:TRect;
    sPaint:PAINTSTRUCT;
    DC:HDC;
    LFont:HFONT;
Begin
    If Msg=WM_PAINT Then
    Begin
        DC:=BeginPaint(hDlg,sPaint);
        GetClientRect(hDlg,sRect);
 //SetTextColor(DC,$A0A0A0);
        SetTextColor(DC,$03FF09); //滚动字幕呈 绿色
        SetBkMode(DC,TRANSPARENT);
        LFont:=CreateFont(-$C,0,0,0,0,0,0,0,1,0,0,0,0,'宋体');
        SelectObject(DC,LFont);
        DrawText(DC,szScroll,-1,sRect,DT_CENTER);
        EndPaint(hDlg,sPaint);
        DeleteObject(LFont);
    End
    Else
    Begin
        CallWindowProc(Pointer(GetWindowLong(hDlg,GWL_USERDATA)),hDlg,Msg,wParam,lParam);
    End;
    Result:=1;
End;

//////////////////////////////////////////////////////////////////
//关于对话框过程函数

Function AboutProc(hDlg:HWND;Msg,wParam,lParam:DWORD):LRESULT;stdcall;
Var
    sRect:TRect;
    sPaint:PAINTSTRUCT;
    sPoint:TPoint;
Begin
    Result:=0;
    Case Msg of
    WM_COMMAND:
        Begin
            Case wParam of
            ABOUT_OK:
                Begin
                    KillTimer(hDlg,$A8);
                    EndDialog(hDlg,0);
                End;
            ABOUT_CLOSE:
                Begin
                    KillTimer(hDlg,$A8);
                    EndDialog(hDlg,0);
                End;
            End;
        End;
    WM_PAINT:
        Begin
            Paint(BeginPaint(hDlg,sPaint),h_Icon,szMainCaption,$767676,0,sRectA);
            EndPaint(hDlg,sPaint);
        End;
    WM_DRAWITEM:
        Begin
            ItemDraw(PDrawItemStruct(lParam));
            Result:=0;
        End;
    WM_INITDIALOG:
        Begin
            GetClientRect(hDlg,sRectA);
            sRectA.Bottom:=sRectA.Top+$14;
            h_ScrollParent:=GetDlgItem(hDlg,$BBD);
            SendMessage(GetDlgItem(hDlg,$BBB),WM_SETFONT,CreateFont(-$C,0,0,0,$2BC,0,0,0,1,0,0,0,0,'宋体'),0);
            SetDlgItemText(hDlg,$BBB,szKeyGenName);
            SetDlgItemText(hDlg,$BBC,szCracker);
            SetWindowText(hDlg,'关于');
            GetClientRect(h_ScrollParent,sRect);
            ScrollWidth:=sRect.right-sRect.left;
            ScrollHeight:=sRect.bottom-sRect.top;
            xCount:=ScrollHeight;
            LineCount:=CountCRLF(szScroll);
            h_Scroll:=CreateWindowEx(WS_EX_LEFT,'Static','',WS_CHILD OR WS_VISIBLE OR SS_CENTER,0,ScrollHeight,ScrollWidth,LineCount*12,h_ScrollParent,0,h_Inst,nil);
            SetWindowLong(h_Scroll,GWL_USERDATA,SetWindowLong(h_Scroll,GWL_WNDPROC,LongWord(@ScrollProc)));
            DialogInit(hDlg);
            SetTimer(hDlg,$A8,60,NIL);
            Result:=1;
        End;
    WM_CTLCOLORDLG:
        Begin
            SetTextColor(wParam,$A0A0A0);
            SetBkMode(wParam,TRANSPARENT);
            Result:=h_Brush;
        End;
    WM_CTLCOLORSTATIC:
        Begin
            //SetTextColor(wParam,$FE248B); //关于对话框 题目 颜色
            SetTextColor(wParam,$FC77B5);
            SetBkMode(wParam,TRANSPARENT);
            Result:=h_Brush;
        End;
    WM_LBUTTONDOWN:
        Begin
            sPoint.x:=lParam AND $FFFF;
            sPoint.y:=lParam SHR 16;
            If PtInRect(sRectA,sPoint) Then
            Begin
                PostMessage(hDlg,WM_NCLBUTTONDOWN,2,0);
            End;
        End;
    WM_TIMER:
        Begin
            Sleep(20);
            Dec(xCount);
            SetWindowPos(h_Scroll,HWND_TOP,0,xCount,ScrollWidth,LineCount*12,0);
            If xCount<(0-LineCount*12) Then
            Begin
                xCount:=ScrollHeight;
            End;
        End;
    End;
End;

//////////////////////////////////////////////////////////////////
//超连接函数

Function LinkProc(hDlg:HWND;Msg,wParam,lParam:DWORD):LRESULT;stdcall;
Begin
    Result:=1;
    Case Msg of
    WM_SETCURSOR:
        Begin
            SetCursor(h_Cur);
        End;
    WM_NCHITTEST:
        Begin
            Result:=1;
        End;
    WM_LBUTTONUP:
        Begin
            ShellExecute(0,nil,'http://17339836.qzone.qq.com',nil,nil,0);//偷了点懒
        End;
    Else
        Begin
            CallWindowProc(Pointer(GetWindowLong(hDlg,GWL_USERDATA)),hDlg,Msg,wParam,lParam);
        End;
    End;
End;

//_____________________________________________________________
//主窗口过程函数

Function MainProc(hDlg:HWND;Msg,wParam,lParam:DWORD):LRESULT;stdcall;
Var
    sPaint:PAINTSTRUCT;
    sPoint:TPoint;
    LFont:HFONT;
    LinkHWND:HWND;
    SetLongRet:Longint;
Begin
    Result:=0;
    
    Case Msg of
    WM_CTLCOLOREDIT:
        Begin
            SetTextColor(wParam,$A0A0A0);
            SetBkMode(wParam,TRANSPARENT);
            Result:=h_Brush;
        End;
    WM_COMMAND:
        Begin
            Case wParam AND $FFFF of
            MAIN_CALC:
                Begin
                    GetDlgItemText(hDlg,$7DA,@RegName,255);
                    //GetRegCode; // 计算注册码过程
                    PatchFile; //打补丁过程
                    SetDlgItemText(hDlg,$7D9,@RegCode);
                End;
            MAIN_EXIT:
                Begin
                    EndDialog(hDlg,0);
                End;
            MAIN_ABOUT:
                Begin
                    MessageBeep(0);
                    DialogBoxParam(h_Inst,LPCTSTR(IDD_ABOUTDLG),hDlg,@AboutProc,0);
                End;
            MAIN_CLOSE:
                Begin
                    EndDialog(hDlg,0);
                End;
            $7DA://就是RegName的ID
                Begin
                    Case wParam SHR 16 of
                    EN_CHANGE: // 在用户名中输入数据时，捕获消息，立即进行计算。过程相当于按下“计算”按键过程。
                        Begin
                            GetDlgItemText(hDlg,$7DA,@RegName,255);
                            //GetRegCode;
                            PatchFile; //打补丁过程;;
                            SetDlgItemText(hDlg,$7D9,@RegCode);
                        End;
                End;
            End;
        End;
    End;
    WM_PAINT:
        Begin
            Paint(BeginPaint(hDlg,sPaint),h_Icon,szMainCaption,$767676,0,sRectM);
            EndPaint(hDlg,sPaint);
        End;
    WM_DRAWITEM:
        Begin
            ItemDraw(PDrawItemStruct(lParam));
            Result:=0;
        End;
    WM_INITDIALOG:
        Begin
            GetClientRect(hDlg,sRectM);
            sRectM.Bottom:=sRectM.Top+$14;
            SetDlgItemText(hDlg,2011,szLink);
            SetDlgItemText(hDlg,2008,szTime);
            SetDlgItemText(hDlg,2005,szName);
            SetDlgItemText(hDlg,2006,szCracker);
            SetDlgItemText(hDlg,2007,szBy);
            SetDlgItemText(hDlg,2012,szCopyright);
            SetWindowText(hDlg,szMainCaption); //任务栏显示名称；
            LFont:=CreateFont(-$C,0,0,0,$2BC,0,1,0,1,0,0,0,0,'宋体');
            LinkHWND:=GetDlgItem(hDlg,2011);
            SendMessage(LinkHWND,WM_SETFONT,LFont,0);//设置字体
            SetLongRet:=SetWindowLong(LinkHWND,GWL_WNDPROC,LongWord(@LinkProc));
            SetWindowLong(LinkHWND,GWL_USERDATA,SetLongRet);
            DialogInit(hDlg);
        End;
    WM_CTLCOLORDLG:
        Begin
            SetTextColor(wParam,$A0A0A0);
            SetBkMode(wParam,TRANSPARENT);
            Result:=h_Brush;
        End;
    WM_CTLCOLORSTATIC:
        Begin
            SetTextColor(wParam,$A0A0A0);
            SetBkMode(wParam,TRANSPARENT);
            Result:=h_Brush;
        End;
    WM_LBUTTONDOWN:
        Begin
            sPoint.x:=lParam AND $FFFF;
            sPoint.y:=lParam SHR 16;
            If PtInRect(sRectM,sPoint) Then
            Begin
                PostMessage(hDlg,WM_NCLBUTTONDOWN,2,0);
            End;
        End;
    WM_CLOSE:
        begin
          EndDialog(hDlg,0);
        end;
    End;
End;

//程序运行时从这里开始执行代码；

begin
    h_Inst:=GetModuleHandle(nil);
    h_Brush:=CreateSolidBrush($0F0F0F);
    h_Cur:=LoadCursor(h_Inst,MAKEINTRESOURCE(LINKCURSOR));
    h_Icon:=LoadIcon(h_Inst,MAKEINTRESOURCE(MAINICON));
    DialogBox(h_Inst,LPCTSTR(IDD_MAINDLG),0,@MainProc); //显示主窗口
    DeleteObject(h_Brush); // 删除窗体
    ExitProcess(0); //退出程序
end.


