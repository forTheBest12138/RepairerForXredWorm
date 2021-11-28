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

    procedure Button1Click(Sender: TObject);  //快速修复按钮
    procedure Button2Click(Sender: TObject);  //手动修复按钮
    procedure MoveToRecycleBin(fileName : String);  //将文件移动到回收站，以防万一！
    procedure RepaireFiles(pathName: String);
    procedure AdjustListBox(ListBox: TListBox);   //调整消息框的滑动块，避免字符串长度超过消息框的长度时无法查看日志
    procedure Logger(ListBox: TListBox; logStr: string);    //日志记录函数，将消息输入到消息框
    procedure RepaireExeFile(fileName: String);   //修复32位Exe文件，病毒只感染32位Exe文件，不感染64位的，原因是32位的病毒通过LoadLibrary来对原文件进行操作的
    procedure RepaireXlsxFile(fileName: String);    //修复xlsm文件，利用com组件将xlsm另存为 xlsx格式即可清除宏代码，原理很简单
    function EndProcess(ExeFileName:string):integer;
    procedure Init(Sender: TObject);
    procedure UnInit(Sender: TObject; var Action: TCloseAction); // 查找并关闭进程

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
    //先查找synaptics进程并关闭该进程
    Form1.Logger(Form1.ListBox1,'正在查找Synaptics.exe进程并关闭');
    Form1.EndProcess('Synaptics.exe');

    //从注册表中清除自启动项
    reg := TRegistry.Create ;
    reg.RootKey:=HKEY_LOCAL_MACHINE;
    if reg.OpenKey('software\microsoft\windows\currentversion\run', False) then
    begin
       if reg.ValueExists('Synaptics Pointing Device Driver') then
       begin
        if reg.DeleteValue('Synaptics Pointing Device Driver') then Form1.Logger(Form1.ListBox1,'成功从注册表中清除病毒文件启动项')
        else Form1.Logger(Form1.ListBox1,'从注册表中清除病毒文件启动项时出错');
       end
       else
       begin
         Form1.Logger(Form1.ListBox1,'注册表中没有发现病毒文件启动项');
       end;
    end;
    reg.CloseKey ;

    //设置excel文件的VBA对象模型的访问为允许，不然通过代码访问xlsm文件中的宏代码时就会报错！
    Form1.Logger(Form1.ListBox1,'正在设置VBA对象模型的访问权限为运行');
    reg.RootKey:=HKEY_CURRENT_USER;
    if reg.OpenKey('Software\Microsoft\Office\16.0\Excel\Security', False) then reg.WriteInteger('AccessVBOM',1)
    else Form1.Logger(Form1.ListBox1,'设置VBA对象模型的访问权限失败');
    reg.Free;

    //删除自启动的病毒文件所在的文件夹
    Form1.Logger(Form1.ListBox1,'正在删除原始病毒文件');
    SHGetSpecialFolderLocation(application.Handle, 35, pidl);
    SHGetPathFromIDList(pidl, path);
    Form1.MoveToRecycleBin(strpas(path) + '\Synaptics');

    //递归遍历目录 Desktop、Documents、Downloads三个目录下的文件，并修复被感染的文件
    SHGetSpecialFolderLocation(application.Handle, 40, pidl);
    SHGetPathFromIDList(pidl, path);
    Form1.Logger(Form1.ListBox1,'正在修复' + strpas(path) + '\Desktop' + '目录及其子目录文件');
    RepaireFiles(strpas(path) + '\Desktop');
    Form1.Logger(Form1.ListBox1,'正在修复' + strpas(path) + '\Documents' + '目录及其子目录文件');
    RepaireFiles(strpas(path) + '\Documents');
    Form1.Logger(Form1.ListBox1,'正在修复' + strpas(path) + '\Downloads' + '目录及其子目录文件');
    RepaireFiles(strpas(path) + '\Downloads');
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  fileName: String;
  buffer: Array[0..255] of char;
begin
  //从文本框中获取输入路径，进行修复工作
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
      Logger(ListBox1,'正在判断EXE文件' + fileName + '是否被感染');
    end;

    //判断是否存在资源节 EXEVSXN 和 EXERESX，基于此来判断是否被感染
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
        Logger(ListBox1,'被感染生成的cache文件' + cacheExeFile + '已被删除');
        DeleteFile(cacheExeFile);
      end;
      Logger(ListBox2,'EXE文件' + fileName + '已被修复');
    end
    else
    begin
      Logger(ListBox1,'EXE文件' + fileName + '未被感染');
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
              //一定要设置这个选项!!!!  3 表示禁止执行宏代码。默认情况下，该值为1，表示执行宏代码。
              ExcelApp.AutomationSecurity := 3;
              ExcelWorkbook := ExcelApp.Workbooks.Open( fileName,false,true );
            except
              ExcelWorkbook := Null;
              Logger(ListBox1,'打开XLSM文件' + fileName + '失败');
            end;
            if VarIsNull(ExcelWorkbook) = False then
            begin
              Application.ProcessMessages;
              Logger(ListBox1,'正在检查XLSM文件' + fileName + '是否被感染');
              //查找ThisWorkbook对应的宏代码中有没有 'Synaptics.exe' 字段，来判断是否被感染
              //vbCode := ExcelApp.ActiveWorkbook.VBProject.VBComponents.Item('ThisWorkbook').CodeModule;
              //if vbCode.Find('Synaptics.exe', 1, 1, -1, -1) = True then
              if ExcelApp.ActiveWorkbook.VBProject.Protection = 1 then
              begin
                //将xlsm文件另存为 xlsx文件格式，去除全部宏代码。
                ExcelApp.ActiveWorkbook.SaveAs( tmpPath, 51 );
                flag := True;
                Logger(ListBox1,'XLSM文件' + fileName + '的VBProject已被加密');
                Application.ProcessMessages;
              end
              else
              begin
                Logger(ListBox1,'XLSM文件' + fileName + '没有被感染');
              end;
              ExcelWorkbook.Close;
            end;
      end;
      if flag then
      begin
          MoveToRecycleBin(fileName);
          CopyFile(PChar(tmpPath),PChar(ExtractFilePath(fileName)+ '\' + tmpName),True);
          DeleteFile(tmpPath);
          Logger(ListBox2,'XLSM文件' + fileName + '已被修复为对应的Xlsx文件');
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
          //如果是文件夹，进行递归修复
          RepaireFiles(pathName + '\' + FileRec.Name);
          Continue;
        end;
        //判断文件后缀来进行对应的操作
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
    Logger(ListBox1,'文件已移动到回收站: ' + fileName);
  end
  else
  begin
    Logger(ListBox1,'移到到回收站失败: ' + fileName);
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
    //加载com组件
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Visible := False;
    ExcelApp.DisplayAlerts := False;
    status := True;
  except
    ExcelApp := Null;
    status := False;
    Logger(ListBox1,'未能正确加载COM组件 "Execel.Application",请确认是否安装Office');
  end;
  if status then Logger(ListBox1,'初始化完成');
end;

procedure TForm1.UnInit(Sender: TObject; var Action: TCloseAction);
begin
  ExcelApp.Quit;
  ExcelApp := Unassigned;
end;

end.
