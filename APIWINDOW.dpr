//ǰЩ�������������˾�Ŀ�����̣�ͬʱҲ��Delphi��д�˸��������������Ĳ�������ԭ������������Դ�ͷ���ʽ���У���ˣ��������ɵĲ������������ޱȣ��������������Ŀ���ƽ��ļ���������������������̵��ƽⲹ��˵�ɣ����ɳ�����ѹ������Ȼ��800K���ҡ������ǵ�����ԭ��ֻ���޸���һ��ƫ������ֵ����˴󶯸ɸ�ذ��ƽ����ļ���������Դ�ļ������и��ǣ�ʵ�����²ߣ���û�и��õİ취�ء��϶��У��������Ǿ������Լ���ƫ������������ģ�������Ͼ���Ϊע����򲹶�����Ļ������ԽС��ȻԽ�ã���ˣ������ô�API��д���������ɵĳ�����Ժ�С��С�����¾���Դ���룬���ɵĳ������54K��ѹ����ֻ��40��K��̫�ɰ�������������һ�����������ʹ��GUI�����ֻʹ�ùؼ��ƽ����Ļ������ɳ������ļ�ֻ��7K.�Ǻǡ���������������ϵĲ���Ŷ��
//******************** ��������Ҫ�޸ĵĳ�������������� ****************
//    ��*���Ի������*��
program apiwindow;
uses Windows,ShellAPI,Messages,System;
const
szMainCaption = '�������ƫ������������V1.0';
szName = '����Ӣ����￼��ϵͳ'; //Ŀ���ƽ��ļ���
szBy = 'futuring@126.com'; //���������
szCracker = '�������'; //������
szCrackeremail = 'mailto:futuring@126.com'; //������email
szLink = 'http://hi.baidu.com/beyond0769'; //��ҳ��ַ
szTime = '2008-03-18'; //�ƽ�ʱ��
szCopyright = '������ֻ���о�ѧϰ��,��ֹ�Ƿ���;';
szKeyGenName = '�������ƫ������������';
szScroll = '�������ƫ������������' + #$0A + #$0A
          + '��ģ���ô�API������д�����湹��'+ #$0A + #$0A
          + '��������Դ�ļ���ʽ������ʡ��'+ #$0A + #$0A
          + '���ɺ���ļ������ѹ����ֻ��'+ #$0A + #$0A
          + '41k���ҡ�ͬʱ������uFMOD��Ԫ' + #$0A + #$0A
          + 'ʵ�ֱ�������ѭ�����š� ' + #$0A + #$0A + #$0A
          + '�ر���л��' + #$0A + #$0A
          + 'you_know[FCG]'+ #$0A+ #$0A
          + '��ѩѧԺ��Delphiţ����' + #$0A+ #$0A
          + '�Լ�uFMOD�ĳ���Ա'+ #$0A+ #$0A
          + '��лһֱ�ڱ���ĬĬ֧���ҵ��ٶ�' + #$0A + #$0A
          + '����ʹ�÷�����'+ #$0A + #$0A
          + '�Ѳ������Ƶ�Ŀ���ļ���ͬĿ¼��'+ #$0A + #$0A
          + '���У������Ӧ�ò��������ɡ� '+ #$0A + #$0A
          + '������ֻ���о�ѧϰ��' + #$0A+#$0A
          + '��ֹһ�зǷ���ҵ��;' + #$0A+#$0A
          + '�����ƽ� by �������(������)' + #$0A+#$0A
          + '������ by �������(������)' + #$0A+#$0A
          + 'futuring@126.com' + #$0A + #$0A
          + 'http://beyond-0769.blog.163.com/' + #$0A + #$0A
          + '2008-03-18';
var
    FileName: PChar = 'SpokenEngMain.exe'; //�ƽ�Ŀ���ļ���������
    IntFileSize: Cardinal = 2731008; //�ƽ�Ŀ���ļ��Ĵ�С�ֽ�
    RBuffer: array[0..1] of Byte = ($75, $0C); //Ŀ���ƽ��ļ�ԭ�е�ƫ����
    WBuffer: array[0..1] of Byte = ($EB, $0C); //�޸ĺ��ƫ����
    OffsetPos: TOVERLAPPED = (Internal: 0; InternalHigh: 0; Offset: $000EE089; OffsetHigh: 0; hEvent: 0);
    CommandLine: PChar; // ���������������ַ��ͨ��W32DSM���Բ鵽
    hwndname: HWND;
    hFile: THANDLE;
    Numb: DWORD;
    Buffer: array[0..1] of Byte;
    nType: DWORD;
    pMsg: string;
////////////////////
//ȫ�ֱ����������� ******************** ���³������������������Ҫ�޸� ******************8
////////////////////
    h_Icon: HICON; //ʵ�����
    h_Inst: HMODULE;
    h_Cur: hCursor;
    h_Brush: HBRUSH;
    sRectL, sRectM, sRectA: TRect;
 //RegName,RegCode:Array[1..256] of Byte; //ԭ�����ݣ��¾����޸ĺ��
    RegName, RegCode: string;
    h_ScrollParent, h_Scroll: HWND;
    ScrollWidth, ScrollHeight: Word;
    xCount: Integer;
    LineCount: Word;
const
////////////////
//��Դ��������//
////////////////
    MAINICON = 100;
    LINKCURSOR = 112;

    IDD_LICENSEDLG = 1000;
    IDD_MAINDLG = 2000;
    IDD_ABOUTDLG = 3000;

    LICENSE_YES = 1001;
    LICENSE_NO = 1002;
    LICENSE_CLOSE = 1004;

    MAIN_CALC = 2001; //���㰴ť�ؼ�ID����ʵ���Ѿ�ȥ����
    MAIN_EXIT = 2002; //�˳���ť�ؼ�ID
    MAIN_ABOUT = 2003; //���ڰ�ť�ؼ�ID
    MAIN_CLOSE = 2004; //�رհ�ť�ؼ�ID

    ABOUT_OK = 3001;
    ABOUT_CLOSE = 3002;

{$R dialog.res}
//++++++++++++++++++�ؼ��ı��ƹ���++++++++++++++++++++
procedure PatchFile;
var
    Res:boolean;
begin
    Setfileattributes(FileName, FILE_ATTRIBUTE_NORMAL + FILE_ATTRIBUTE_ARCHIVE); //�����ļ�������Ϊ����
    hFile := CreateFile(FileName, GENERIC_READ or GENERIC_WRITE, FILE_SHARE_READ or
    FILE_SHARE_WRITE, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
    try
        if hFile <> INVALID_HANDLE_VALUE then
        begin
            if GetFileSize(hFile,nil)<>IntFileSize then
            begin
                nType := MB_OK or MB_ICONINFORMATION;
                pMsg := '�ļ�У�����'+#13+'�ļ���С��һ�£�Ŀ���ļ�(Ӣ�����������:' + FileName + ')�����Ѿ����޸ģ��ܾ��򲹶���';
                exit;
            end;
            ReadFile(hFile, Buffer, 2, Numb, @OffsetPos); //���������ƫ���������Ѿ����ϲ��������ݣ�����ʾ�Ѿ����ϲ�����
            if Word(Buffer[0]) = Word(WBuffer[0]) then begin
                nType := MB_OK or MB_ICONINFORMATION;
                pMsg := '������ʾ��'+#13+'Ŀ�����(' + FileName + ')�Ѿ���������ˣ����������� ^_^';
                Res:=True;
                exit;
                end;
            if Word(Buffer[0]) = Word(RBuffer[0]) then {// ��ȡƫ���Ƿ���ȷ��} begin
                CopyFile(FileName, PChar('����' + FileName), False); //�����ƽ�Ŀ���ļ���
                if WriteFile(hFile, WBuffer, 2, Numb, @OffsetPos) then
                begin
                    nType := MB_OK or MB_ICONINFORMATION;
                    pMsg := '�ɹ����ϲ���! Enjoy it^_^' + #13 +#13+
                            'Cracked & Coded by �������(������)'; ; //�ɹ�������
                    Res:=True;
                end else
                begin // д�����
                    nType := MB_OK or MB_ICONERROR;
                    pMsg := '������ʾ��'+#13+'������Ŀ���ļ�(' + FileName + ')�Ƿ��Ѿ����򿪣����Ƚ���ر��ٳ��ԡ�';
                end;
            end else {//ƫ�Ʋ���ȷ����ʾ����}
            begin
                nType := MB_OK or MB_ICONWARNING;
                pMsg := '������ʾ��'+#13+'Ŀ���ļ�(' + FileName + ')ƫ�����������ܾ��򲹶���'+#13+'���������뵽������ҳ��ӳ��';
                ShellExecute(0, nil, szLink, nil, nil, 0);
            end;
        end else {//�ļ����򿪻��ļ��Ҳ���}
        begin
            nType := MB_OK or MB_ICONERROR;
            pMsg := '�������� :-( ����ԭ�����£�' + #13 +
                    '1.Ŀ���ļ��Ҳ���:(Ӣ�����������:' + FileName + ')����ѱ������Ƶ�ͬһ��Ŀ¼�������С�' + #13 +
                    '2.Ŀ���ļ��Ѿ����򿪣����Ƚ���ر��ٳ��ԡ�';
        end;
    finally
        CloseHandle(hFile);
        MessageBox(0, PChar(pMsg), PChar('���������ܰ��ʾ��'), nType);
        if Res then ShellExecute(0, nil, szLink, nil, nil, 0);
    end;
end;


//+++++++++++++++++����Ϊ�����������������ĺ�������++++++++++++++

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
    //SetTextColor(lpDIS^.hDC,$03FF09); //��ɫ
    SetBkMode(lpDIS^.hDC,TRANSPARENT);
    DrawEdge(lpDIS^.hDC,lpDIS^.rcItem,BDR_RAISEDOUTER,BF_RECT);
    GetWindowText(lpDIS^.hwndItem,@Str,10);
    DrawText(lpDIS^.hDC,@Str,-1,lpDIS^.rcItem,DT_SINGLELINE OR DT_VCENTER OR DT_CENTER);

    If lpDIS^.itemState MOD 2=1 Then
    Begin
        SetTextColor(lpDIS^.hDC,$03FF09); //���� ��ɫ
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
    LogBrush.lbColor:=$9E6A54; //��������ɫ
    LhBR:=CreateBrushIndirect(LogBrush);
    FrameRect(dc,Rect,LhBR);
    DeleteObject(LhBR);
    SetTextColor(dc,$05E8FC);
    SetBkMode(dc,TRANSPARENT);
    Rect.left:=2;
    Rect.top:=2;
    Dec(Rect.bottom,2);
    LFont:=CreateFont(-$C,0,0,0,$2BC,0,0,0,1,0,0,0,0,'����');
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

//�����ı�����

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
        SetTextColor(DC,$03FF09); //������Ļ�� ��ɫ
        SetBkMode(DC,TRANSPARENT);
        LFont:=CreateFont(-$C,0,0,0,0,0,0,0,1,0,0,0,0,'����');
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
//���ڶԻ�����̺���

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
            SendMessage(GetDlgItem(hDlg,$BBB),WM_SETFONT,CreateFont(-$C,0,0,0,$2BC,0,0,0,1,0,0,0,0,'����'),0);
            SetDlgItemText(hDlg,$BBB,szKeyGenName);
            SetDlgItemText(hDlg,$BBC,szCracker);
            SetWindowText(hDlg,'����');
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
            //SetTextColor(wParam,$FE248B); //���ڶԻ��� ��Ŀ ��ɫ
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
//�����Ӻ���

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
            ShellExecute(0,nil,'http://17339836.qzone.qq.com',nil,nil,0);//͵�˵���
        End;
    Else
        Begin
            CallWindowProc(Pointer(GetWindowLong(hDlg,GWL_USERDATA)),hDlg,Msg,wParam,lParam);
        End;
    End;
End;

//_____________________________________________________________
//�����ڹ��̺���

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
                    //GetRegCode; // ����ע�������
                    PatchFile; //�򲹶�����
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
            $7DA://����RegName��ID
                Begin
                    Case wParam SHR 16 of
                    EN_CHANGE: // ���û�������������ʱ��������Ϣ���������м��㡣�����൱�ڰ��¡����㡱�������̡�
                        Begin
                            GetDlgItemText(hDlg,$7DA,@RegName,255);
                            //GetRegCode;
                            PatchFile; //�򲹶�����;;
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
            SetWindowText(hDlg,szMainCaption); //��������ʾ���ƣ�
            LFont:=CreateFont(-$C,0,0,0,$2BC,0,1,0,1,0,0,0,0,'����');
            LinkHWND:=GetDlgItem(hDlg,2011);
            SendMessage(LinkHWND,WM_SETFONT,LFont,0);//��������
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

//��������ʱ�����￪ʼִ�д��룻

begin
    h_Inst:=GetModuleHandle(nil);
    h_Brush:=CreateSolidBrush($0F0F0F);
    h_Cur:=LoadCursor(h_Inst,MAKEINTRESOURCE(LINKCURSOR));
    h_Icon:=LoadIcon(h_Inst,MAKEINTRESOURCE(MAINICON));
    DialogBox(h_Inst,LPCTSTR(IDD_MAINDLG),0,@MainProc); //��ʾ������
    DeleteObject(h_Brush); // ɾ������
    ExitProcess(0); //�˳�����
end.


