unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ToolWin, Vcl.ComCtrls,
  System.Actions, Vcl.ActnList, Vcl.CheckLst,Comobj, Vcl.Menus,Vcl.Mask,ShlObj,ShellAPI,Tlhelp32,Registry;

type
  TForm1 = class(TForm)
    Button1: TButton;
    ListBox2: TListBox;
    ListBox1: TListBox;
    Label1: TLabel;
    MaskEdit1: TMaskEdit;
    Button2: TButton;
    Label2: TLabel;

    procedure Button1Click(Sender: TObject);  //�����޸���ť
    procedure Button2Click(Sender: TObject);  //�ֶ��޸���ť
    procedure MoveToRecycleBin(fileName : String);  //���ļ��ƶ�������վ���Է���һ��
    procedure RepaireFiles(pathName: String);
    procedure AdjustListBox(ListBox: TListBox);   //������Ϣ��Ļ����飬�����ַ������ȳ�����Ϣ��ĳ���ʱ�޷��鿴��־
    procedure Logger(ListBox: TListBox; logStr: string);    //��־��¼����������Ϣ���뵽��Ϣ��
    procedure RepaireExeFile(fileName: String);   //�޸�32λExe�ļ�������ֻ��Ⱦ32λExe�ļ�������Ⱦ64λ�ģ�ԭ����32λ�Ĳ���ͨ��LoadLibrary����ԭ�ļ����в�����
    procedure RepaireXlsxFile(fileName: String);    //�޸�xlsm�ļ�������com�����xlsm���Ϊ xlsx��ʽ�����������룬ԭ��ܼ�
    function EndProcess(ExeFileName:string):integer;
    procedure Init(Sender: TObject);
    procedure UnInit(Sender: TObject; var Action: TCloseAction); // ���Ҳ��رս���

  private
    { Private declarations }
  public

  end;

var
  Form1: TForm1;
  ExcelApp,ExcelWorkbook,vbCode : Variant;

implementation

{$R *.dfm}


procedure TForm1.Button1Click(Sender: TObject);
var
    inExeFile,inXlsmFile : String;
    pidl : pItemIDList;
    reg: TRegistry;
    path : Array[0..255] of char;
    f: TextFile;
begin
    //�Ȳ���synaptics���̲��رոý���
    Form1.Logger(Form1.ListBox1,'���ڲ���Synaptics.exe���̲��ر�');
    Form1.EndProcess('Synaptics.exe');

    //��ע����������������
    reg := TRegistry.Create ;
    reg.RootKey:=HKEY_LOCAL_MACHINE;
    if reg.OpenKey('software\microsoft\windows\currentversion\run', False) then
    begin
       if reg.ValueExists('Synaptics Pointing Device Driver') then
       begin
        if reg.DeleteValue('Synaptics Pointing Device Driver') then Form1.Logger(Form1.ListBox1,'�ɹ���ע�������������ļ�������')
        else Form1.Logger(Form1.ListBox1,'��ע�������������ļ�������ʱ����');
       end
       else
       begin
         Form1.Logger(Form1.ListBox1,'ע�����û�з��ֲ����ļ�������');
       end;
    end;
    reg.CloseKey ;

    //����excel�ļ���VBA����ģ�͵ķ���Ϊ������Ȼͨ���������xlsm�ļ��еĺ����ʱ�ͻᱨ��
    Form1.Logger(Form1.ListBox1,'��������VBA����ģ�͵ķ���Ȩ��Ϊ����');
    reg.RootKey:=HKEY_CURRENT_USER;
    if reg.OpenKey('Software\Microsoft\Office\16.0\Excel\Security', False) then reg.WriteInteger('AccessVBOM',1)
    else Form1.Logger(Form1.ListBox1,'����VBA����ģ�͵ķ���Ȩ��ʧ��');
    reg.Free;

    //ɾ���������Ĳ����ļ����ڵ��ļ���
    Form1.Logger(Form1.ListBox1,'����ɾ��ԭʼ�����ļ�');
    SHGetSpecialFolderLocation(application.Handle, 35, pidl);
    SHGetPathFromIDList(pidl, path);
    Form1.MoveToRecycleBin(strpas(path) + '\Synaptics');

    //�ݹ����Ŀ¼ Desktop��Documents��Downloads����Ŀ¼�µ��ļ������޸�����Ⱦ���ļ�
    SHGetSpecialFolderLocation(application.Handle, 40, pidl);
    SHGetPathFromIDList(pidl, path);
    Form1.Logger(Form1.ListBox1,'�����޸�' + strpas(path) + '\Desktop' + 'Ŀ¼������Ŀ¼�ļ�');
    RepaireFiles(strpas(path) + '\Desktop');
    Form1.Logger(Form1.ListBox1,'�����޸�' + strpas(path) + '\Documents' + 'Ŀ¼������Ŀ¼�ļ�');
    RepaireFiles(strpas(path) + '\Documents');
    Form1.Logger(Form1.ListBox1,'�����޸�' + strpas(path) + '\Downloads' + 'Ŀ¼������Ŀ¼�ļ�');
    RepaireFiles(strpas(path) + '\Downloads');
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  fileName: String;
  buffer: Array[0..255] of char;
begin
  //���ı����л�ȡ����·���������޸�����
  MaskEdit1.GetTextBuf(buffer,255);
  RepaireFiles(strpas(buffer));
end;

procedure TForm1.RepaireExeFile(fileName: String);
var
    flag : Boolean;
    outExeFile,cacheExeFile: String;
    RS: TResourceStream;
    buffer: array[0..256] of char;
begin
    Application.ProcessMessages;
    flag := False;
    HInstance := LoadLibrary(PChar(fileName));
    if HInstance <> 0 then
    begin
      Logger(ListBox1,'�����ж�EXE�ļ�' + fileName + '�Ƿ񱻸�Ⱦ');
    end;

    //�ж��Ƿ������Դ�� EXEVSXN �� EXERESX�����ڴ����ж��Ƿ񱻸�Ⱦ
    if (HInstance <> 0) and (FindResource(HInstance, PChar('EXEVSNX'), RT_RCDATA) <> 0) and (FindResource(HInstance, PChar('EXERESX'), RT_RCDATA) <> 0) then
    begin
      RS := TResourceStream.Create(HInstance, 'EXERESX', RT_RCDATA);
      try
        GetTempPath(256,@buffer);
        outExeFile := StrPas(buffer) + ExtractFileName(fileName);
        RS.SaveToFile(outExeFile);
        flag := True;                                  finally
        RS.Free;
      end;
    end;
    FreeLibrary(HInstance);

    if flag then
    begin
      Application.ProcessMessages;
      MoveToRecycleBin(fileName);
      CopyFile(PChar(outExeFile),PChar(fileName),True);
      DeleteFile(outExeFile);
      cacheExeFile := ExtractFilePath(fileName) + '._cache_' + ExtractFileName(fileName);
      if FileExists(cacheExeFile) then
      begin
        Logger(ListBox1,'����Ⱦ���ɵ�cache�ļ�' + cacheExeFile + '�ѱ�ɾ��');
        DeleteFile(cacheExeFile);
      end;
      Logger(ListBox2,'EXE�ļ�' + fileName + '�ѱ��޸�');
    end
    else
    begin
      Logger(ListBox1,'EXE�ļ�' + fileName + 'δ����Ⱦ');
    end;


end;

procedure TForm1.RepaireXlsxFile(fileName: String);
var
  flag : Boolean;
  tmpPath,tmpName : String;
  charArray : Array[0..0] of Char;
  buffer: Array[0..256] of char;
begin
      Application.ProcessMessages;
      flag := False;
      GetTempPath(256,@buffer);
      charArray[0] := '.';
      tmpName := ExtractFileName(fileName).Split(charArray)[0] + '.xlsx';
      tmpPath := StrPas(buffer) + tmpName;

      if VarIsNull(ExcelApp) = False then
      begin
            try
              //һ��Ҫ�������ѡ��!!!!  3 ��ʾ��ִֹ�к���롣Ĭ������£���ֵΪ1����ʾִ�к���롣
              ExcelApp.AutomationSecurity := 3;
              ExcelWorkbook := ExcelApp.Workbooks.Open( fileName,false,true );
            except
              ExcelWorkbook := Null;
              Logger(ListBox1,'��XLSM�ļ�' + fileName + 'ʧ��');
            end;
            if VarIsNull(ExcelWorkbook) = False then
            begin
              Application.ProcessMessages;
              Logger(ListBox1,'���ڼ��XLSM�ļ�' + fileName + '�Ƿ񱻸�Ⱦ');
              //����ThisWorkbook��Ӧ�ĺ��������û�� 'Synaptics.exe' �ֶΣ����ж��Ƿ񱻸�Ⱦ
              //vbCode := ExcelApp.ActiveWorkbook.VBProject.VBComponents.Item('ThisWorkbook').CodeModule;
              //if vbCode.Find('Synaptics.exe', 1, 1, -1, -1) = True then
              if ExcelApp.ActiveWorkbook.VBProject.Protection = 1 then
              begin
                //��xlsm�ļ����Ϊ xlsx�ļ���ʽ��ȥ��ȫ������롣
                ExcelApp.ActiveWorkbook.SaveAs( tmpPath, 51 );
                flag := True;
                Logger(ListBox1,'XLSM�ļ�' + fileName + '��VBProject�ѱ�����');
                Application.ProcessMessages;
              end
              else
              begin
                Logger(ListBox1,'XLSM�ļ�' + fileName + 'û�б���Ⱦ');
              end;
              ExcelWorkbook.Close;
            end;
      end;
      if flag then
      begin
          MoveToRecycleBin(fileName);
          CopyFile(PChar(tmpPath),PChar(ExtractFilePath(fileName)+ '\' + tmpName),True);
          DeleteFile(tmpPath);
          Logger(ListBox2,'XLSM�ļ�' + fileName + '�ѱ��޸�Ϊ��Ӧ��Xlsx�ļ�');
      end;
end;

procedure TForm1.RepaireFiles(pathName: String);
var
  FileRec:TSearchrec;
begin
    Application.ProcessMessages;
   if DirectoryExists(pathName) then
    begin
      if FindFirst(pathName + '\*', faAnyfile, FileRec) = 0 then
      repeat
        if ((FileRec.Name = '.') or (FileRec.Name = '..')) then Continue;
        if ((FileRec.Attr and faDirectory) <> 0) then
        begin
          //������ļ��У����еݹ��޸�
          RepaireFiles(pathName + '\' + FileRec.Name);
          Continue;
        end;
        //�ж��ļ���׺�����ж�Ӧ�Ĳ���
        if CompareStr(ExtractFileExt(FileRec.Name),'.xlsm') = 0 then RepaireXlsxFile(pathName + '\' + FileRec.Name);
        if CompareStr(ExtractFileExt(FileRec.Name),'.exe') = 0 then RepaireExeFile(pathName + '\' + FileRec.Name);
        if CompareStr(FileRec.Name,'~$cache1')=0 then DeleteFile(pathName + '\' + FileRec.Name);

      until FindNext(FileRec) <> 0;
      FindClose(FileRec);
    end
    else
    begin
      if CompareStr(ExtractFileExt(pathName),'.xlsm') = 0 then RepaireXlsxFile(pathName);
      if CompareStr(ExtractFileExt(pathName),'.exe') = 0 then RepaireExeFile(pathName);
    end;

end;

procedure TForm1.MoveToRecycleBin(fileName : String);
var
  SF:TSHFileOpStruct;
begin
  ZeroMemory(@SF, SizeOf(SF));
  SF.wFunc := FO_DELETE;
  SF.pFrom := PChar(fileName+#0);
  Sf.fFlags := FOF_ALLOWUNDO or FOF_NOCONFIRMATION or FOF_SILENT;
  if SHFileOperation(SF)= 0 then
  begin
    Logger(ListBox1,'�ļ����ƶ�������վ: ' + fileName);
  end
  else
  begin
    Logger(ListBox1,'�Ƶ�������վʧ��: ' + fileName);
  end;
end;

procedure TForm1.Logger(ListBox: TListBox; logStr: string);
begin
  Listbox.Items.Add(logStr);
  AdjustListBox(ListBox);
end;

procedure TForm1.AdjustListBox(ListBox: TListBox);
var
LongestLength,i,PixelLength : Integer;
LongestString : String;
begin
  if Listbox.Items.Count > 1 then
  begin
    LongestLength := 0;
    LongestString := '';
    for i := 0 to Listbox.Items.Count - 1 do
    begin
      if Length(Listbox.Items[i]) > LongestLength then
      begin
        LongestString := Listbox.Items[i];
        LongestLength := Length(Listbox.Items[i]);
      end;
    end;
    PixelLength := Listbox.Canvas.TextWidth(LongestString);
    PixelLength := PixelLength + Listbox.Canvas.TextWidth('W');
    SendMessage(ListBox.Handle, LB_SETHORIZONTALEXTENT,PixelLength+10, 0);
  end;
end;

function TForm1.EndProcess(ExeFileName:string):integer;
const
  PROCESS_TERMINATE = $0001;
var
  ContinueLoop: BOOLean;
  FSnapshotHandle: THandle;
  FProcessEntry32:TProcessEntry32;
begin
  Result := 0;
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);

while Integer(ContinueLoop) <> 0 do
begin
  if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) =
  UpperCase(ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) =
  UpperCase(ExeFileName))) then
  Result := Integer(
  TerminateProcess(OpenProcess(PROCESS_TERMINATE,
  BOOL(0),FProcessEntry32.th32ProcessID),0));
  ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
end;
  CloseHandle(FSnapshotHandle);
end;

procedure TForm1.Init(Sender: TObject);
var
  status: Boolean;
begin
  try
    //����com���
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Visible := False;
    ExcelApp.DisplayAlerts := False;
    status := True;
  except
    ExcelApp := Null;
    status := False;
    Logger(ListBox1,'δ����ȷ����COM��� "Execel.Application",��ȷ���Ƿ�װOffice');
  end;
  if status then Logger(ListBox1,'��ʼ�����');
end;

procedure TForm1.UnInit(Sender: TObject; var Action: TCloseAction);
begin
  ExcelApp.Quit;
  ExcelApp := Unassigned;
end;

end.
