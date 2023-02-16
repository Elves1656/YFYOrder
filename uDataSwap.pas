unit uDataSwap;//123

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom,
  dxSkinDarkSide, dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle,
  dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast, dxSkiniMaginary,
  dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin,
  dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinPumpkin,
  dxSkinSeven, dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus,
  dxSkinSilver, dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008,
  dxSkinTheAsphaltWorld, dxSkinsDefaultPainters, dxSkinValentine,
  dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue, dxSkinsdxBarPainter,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  dxSkinscxPCPainter, cxContainer, cxEdit, dxLayoutcxEditAdapters,
  cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage, cxNavigator,
  Data.DB, cxDBData, Vcl.StdCtrls, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid,
  cxSplitter, dxLayoutContainer, cxTextEdit, dxLayoutControl, Vcl.ExtCtrls,
  dxBar, DBAccess, Uni, MemDS, Datasnap.DBClient, StrUtils,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdFTP, DateUtils, UniProvider,
  OracleUniProvider, AnsiStrings, Crypt2, IdCoderMIME,
  dxLayoutControlAdapters, System.NetEncoding, IniFiles, IdMessageClient,
  IdSMTPBase, IdSMTP, IdMessage, Data.Win.ADODB, Vcl.OleServer,
  AddressParsing_TLB, CatAddrII, HCT_Library_Addr_TLB, uGoPChomeAddress;

type
  TfmDataSwap = class(TForm)
    dxBarManager1: TdxBarManager;
    dxBarManager1Bar1: TdxBar;
    lbQuery: TdxBarLargeButton;
    Panel1: TPanel;
    Panel2: TPanel;
    dxLayoutControl1: TdxLayoutControl;
    dxLayoutControl1Group_Root: TdxLayoutGroup;
    edtPCKID: TcxTextEdit;
    dxLayoutControl1Item1: TdxLayoutItem;
    upLoad: TdxBarLargeButton;
    downLoad: TdxBarLargeButton;
    lbSingle: TdxBarLargeButton;
    Panel3: TPanel;
    Panel4: TPanel;
    cxSplitter1: TcxSplitter;
    cxGrid2: TcxGrid;
    tv1_: TcxGridDBTableView;
    tv1_Column10: TcxGridDBColumn;
    tv1_Column1: TcxGridDBColumn;
    tv1_Column13: TcxGridDBColumn;
    tv1_Column2: TcxGridDBColumn;
    tv1_Column3: TcxGridDBColumn;
    tv1_Column9: TcxGridDBColumn;
    tv1_Column5: TcxGridDBColumn;
    cxGridLevel2: TcxGridLevel;
    mLog: TMemo;
    btnSave: TButton;
    btnClear: TButton;
    cdsQuery: TUniQuery;
    udsQuery: TUniDataSource;
    cdsTmp: TUniQuery;
    IdFTP1: TIdFTP;
    SaveDialog1: TSaveDialog;
    OpenDialog1: TOpenDialog;
    dbUniECDB: TUniConnection;
    dbUniSR3: TUniConnection;
    cdsCHK: TUniQuery;
    Button1: TButton;
    dxLayoutControl1Item2: TdxLayoutItem;
    IdSMTP1: TIdSMTP;
    IdMessage1: TIdMessage;
    CatAddrII1: TCatAddrII;
    ADOQ_GetZipMapping: TADOQuery;
    Button2: TButton;
    dxLayoutControl1Item3: TdxLayoutItem;
    Edit1: TEdit;
    dxLayoutControl1Item4: TdxLayoutItem;
    Button3: TButton;
    dxLayoutControl1Item5: TdxLayoutItem;
    procedure FormCreate(Sender: TObject);
    procedure lbQueryClick(Sender: TObject);
    procedure upLoadClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure downLoadClick(Sender: TObject);
    procedure lbSingleClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Zoneid, p_KEY: String;
    customerID: String;
    pickNo: TStringList;
    HostIP, ftpUser, ftpPW: String;
    ftpPort: Integer;
    bCatInit: Boolean;
    FCat5ZipCodeVersion: String;
    function OpenSQL(SQL: UTF8String; cds: TUniQuery; act: Integer; flag: Integer): String;
    function CallSP(SPName: String; params, vOut: Variant; flag: Integer): OleVariant;
    function UpLoadFile(fileName: String): Boolean;
    function EnCrypt(CryptText: String): String;
    function DeCrypt(CryptText: String): string;
    function StrToDt(strDT: String): TDateTime;
    function TextFmt(Str: string; iLen: Integer): String;
    function GetTableYear(dDate: TDate; iIsB4: Integer = 0): String;
    function GetCatFiveCode(sType, sOddrecaddr: String): String;
    function Get0003ZipCode(sOddrecaddr: String): String;
    function Get0044ZipCode(sOddrecaddr: String; sType: Char): String;  //2022.09.12 add
    procedure UpdateSND(vfileName: String);
    procedure SendSMS(TelNo, MsgTxt: String);
    procedure SendMail(vType, MsgTxt: String);
    procedure UpdatePCK;
    procedure CheckData;
    procedure StringSaveToFile(AString,AFileName: String);
    function GetCodPay(pNo, ordNo: String): integer;
  end;

const
  {$ifdef debug}
  upDir: String = 'PCH2P_TEST';
  downDir: String = 'P2PCH_TEST';
  {$endif}
  {$ifdef release}
  upDir: String = 'PCH2P';
  downDir: String = 'P2PCH';
  {$endif}
  fComp_Name: String = 'ECAP2';
  UserID: String = 'US00015799';
  UserName: String = '外倉出貨';
  {$ifdef debug}
  //p_KEY = '61D9362544541411B0479C3FD32ADB70E2EA86F4088754C2C07ED722305A4166';
  p_IV = 'A5140D25DBC5E1EF0B134CAB5C70D1F4';
  {$endif}
  //正式機
  {$ifdef release}
  //p_KEY = 'AFB480D31D4E916007E934A754A1C5B28409CC0D98BB5B1885FCF00751DA2455';
  p_IV = '2B455DDF7AC4C9D9D3D3DD0BF4D5E3A4';
  {$endif}

  CAT_CONNECTION_STRING = 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=%s;Persist Security Info=True';
var
  fmDataSwap: TfmDataSwap;
  ac_003: Taddr_Compare2;

implementation

{$R *.dfm}

procedure TfmDataSwap.btnClearClick(Sender: TObject);
begin
  mLog.Lines.Clear;
end;

procedure TfmDataSwap.btnSaveClick(Sender: TObject);
var
  fileName: String;
begin
  try
    SaveDialog1.Filter     := 'Txt files (*.txt)|*.TXT';
    SaveDialog1.FilterIndex:= 1;

    if SaveDialog1.Execute = False then Exit;

    fileName:= SaveDialog1.FileName;

    if Pos('.TXT', UpperCase(fileName)) = 0 then fileName:= fileName + '.txt';

    if (not FileExists(fileName)) or
      (MessageBox(Self.Handle, '檔案已存在, 是否覆蓋？', '儲存檔案', MB_YESNO + MB_APPLMODAL + MB_ICONQUESTION) = IDYES) then
    begin
      mLog.Lines.SaveToFile(fileName);
    end;

  except
    on E: Exception do
      begin
        MessageBox(Self.Handle, PChar(E.Message), '儲存檔案', MB_OK + MB_APPLMODAL + MB_ICONERROR);
        Exit;
      end;
  end;
end;

procedure TfmDataSwap.Button1Click(Sender: TObject);
var
  addList: TStringList;
begin
//

  addlist := TStringList.Create;
  try
    if OpenDialog1.Execute then
      addlist.LoadFromFile(OpenDialog1.FileName);
    addList.Text := EnCrypt(addList.Text);    //加密
    addList.SaveToFile(ExtractFilePath(Application.ExeName) + FormatDateTime('yyyymmddHHNNSS', Now) + '.txt');

  finally
    FreeAndNil(addlist);
  end;
end;

procedure TfmDataSwap.Button2Click(Sender: TObject);
var
  zCode: String;
begin
  //zCode := Get0003ZipCode(edit1.Text);
  zCode := Get0044ZipCode(edit1.Text, '1');
  showmessage(zCode);
end;

procedure TfmDataSwap.Button3Click(Sender: TObject);
var
  sSQL: string;
  fileList: TStringList;
  i: integer;
  vDt: TDateTime;
begin
  //2022.07.05 記Log

  if OpenDialog1.Execute then
  begin
    try
      fileList := TStringList.Create;
      fileList.LoadFromFile(OpenDialog1.FileName);
      for i := 0 to fileList.Count - 1 do
      begin
        try
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 開始更新記錄檔... ');
          sSQL := ' select Max(outdt) outdt from (' +
                  ' select pickrealshipdt_d || '' '' || pickrealshiptm_d outdt from ecoper.pick_send_d1_' + GetTableYear(Date) + ' where pickno_d = ''' + Trim(Copy(fileList[i], 11, 23)) + ''' ' +
                  ' union ' +
                  ' select pickrealshipdt_d || '' '' || pickrealshiptm_d outdt from ecoper.pick_send_d1_' + GetTableYear(Date, 1) + ' where pickno_d = ''' + Trim(Copy(fileList[i], 11, 23)) + ''' ' +
                  ' ) having Max(outdt) is not null ';
          OpenSQL(sSQL, cdsTmp, 1, 1);

          if cdsTmp.RecordCount = 0 then continue;

          vDT := StrToDateTime(cdsTmp.FieldByName('outdt').AsString);
          sSQL := 'select outdt from ecoper.exwarehouse_time where pickid = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ';
          OpenSQL(sSQL, cdsTmp, 1, 1);
          if cdsTmp.RecordCount = 0 then
          begin
            sSQL := 'insert into ecoper.exwarehouse_time (pickid, outdt, keyidt, filename) ' +
                      '  values (''' + Trim(Copy(fileList[i], 11, 20)) + ''', to_date(''' + Trim(Copy(fileList[i], 92, 14)) + ''', ''yyyymmddhh24miss''), to_date(''' + FormatDateTime('yyyy/mm/dd HH:NN:SS', vDT) + ''', ''yyyy/mm/dd hh24:mi:ss''), ''' + OpenDialog1.FileName + ''') ';
            OpenSQL(sSQL, cdsTmp, 2, 1);
          end
          else if cdsTmp.FieldByName('outdt').AsDateTime > StrToDT(Trim(Copy(fileList[i], 92, 14))) then
          begin
            sSQL := 'update ecoper.exwarehouse_time set outdt = to_date(''' + Trim(Copy(fileList[i], 92, 14)) + ''', ''yyyymmddhh24miss'') ' +
                      '  where pickid = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ';
            OpenSQL(sSQL, cdsTmp, 2, 1);
          end;
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 更新記錄檔成功... ');
        except
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 更新記錄檔失敗... ');
        end;
      end;
    finally
      FreeAndNil(fileList);
    end;
  end;
end;

function TfmDataSwap.CallSP(SPName: String; params, vOut: Variant;
  flag: Integer): OleVariant;
var
  UniSP: TUniStoredProc;
  ISfunc: Boolean;
  i, j, k: Integer;
  aOut: array of Variant;
begin
  try

    UniSP := TUniStoredProc.Create(nil);

    if flag = 1 then
    begin
      UniSP.Connection := dbUniECDB;
      try
         try
            if not dbUniECDB.Connected then dbUniECDB.Connected:= True;
         except

         end;
      finally
         dbUniECDB.Connected:= True;
      end;
    end;
    if flag = 2 then
    begin
      UniSP.Connection := dbUniSR3;
      try
         try
            if not dbUniSR3.Connected then dbUniSR3.Connected:= True;
         except

         end;
      finally
         dbUniSR3.Connected:= True;
      end;

    end;

    UniSP.StoredProcName := SPName;
    UniSP.Prepare;
    for i := 0 to UniSP.Params.Count - 1 do
    begin
      Case UniSP.Params[i].ParamType of
        ptInput:
          begin
            UniSP.Params[i].Value := params[i];
          end;
        ptOutput:
          begin
          end;
        ptResult:
          begin
            ISfunc := True;
          end;
      end;
    end;
    UniSP.ExecProc;

    if VarIsArray(vOut) then
    begin
      SetLength(aOut, VarArrayHighBound(vOut, 1) + 1);
      for i := 0 to VarArrayHighBound(vOut, 1) do
      begin
        aOut[i] := UniSP.ParamByName(vOut[i]).Value;
      end;
      result := aOut;
    end;

    //result := aOut;

    //UniSP.Connection.Connected := False;
    FreeAndNil(UniSP);
  except
    on eException: Exception do
    begin
      mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Call SP Error > ' + eException.Message);
    end;
  end;
end;

procedure TfmDataSwap.CheckData;
var
  i: Integer;
  sSQL: String;
begin
  if (FormatDateTime('dddd', now) = '星期六') or (FormatDateTime('dddd', now) = '星期日') then
    Exit;
  if (FormatDateTime('hh', now) < '09') or (FormatDateTime('hh', now) >= '21') then
    Exit;
  sSQL := 'select distinct o.ospickno_m ' +
          '  from ecoper.outstock_m1_' + GetTableYear(Date) + ' o ' +
          'where o.oszoneid_m = ''' + Zoneid + '''  ' +
          '  and o.osstatus_m <> ''SND'' ' +
          '  and o.ossurkind_m = ''OTS'' ' +
          '  and nvl(o.oscp_m, '' '') = '' ''  ' +
          '  and trunc((sysdate - to_date(o.ospickdt_m || o.ospicktm_m, ''YYYY/MM/DD HH24:mi:ss'')) * 24 * 60) > 480 ' +
          ' union all ' +
          'select distinct o.ospickno_m ' +
          '  from ecoper.outstock_m1_' + GetTableYear(Date, 1) + ' o ' +
          'where o.oszoneid_m = ''' + Zoneid + '''  ' +
          '  and o.osstatus_m <> ''SND'' ' +
          '  and o.ossurkind_m = ''OTS'' ' +
          '  and nvl(o.oscp_m, '' '') = '' ''  ' +
          '  and trunc((sysdate - to_date(o.ospickdt_m || o.ospicktm_m, ''YYYY/MM/DD HH24:mi:ss'')) * 24 * 60) > 480 ';

  //OpenSQL(sSQL, cdsCHK, 1, 2);
  OpenSQL(sSQL, cdsCHK, 1, 1);
  if cdsCHK.RecordCount > 0 then
  begin
    SendSMS('0930732789', '宅配通' + Zoneid + ' 庫有( ' + IntToStr(cdsCHK.RecordCount) + ' )筆揀貨單超過8小時已派單未出貨!');
  end;
end;

function TfmDataSwap.DeCrypt(CryptText: String): string;
var
  crypt: HCkCrypt2;
  success: Boolean;
  IvHex: PWideChar;
  KeyHex: PWideChar;
  decStr: PWideChar;
  ADec: TIdDecoderMIME;
begin
  crypt := CkCrypt2_Create();
  ADec := TIdDecoderMIME.Create(Nil);
  try
    success := CkCrypt2_UnlockComponent(crypt, 'MISHIHCrypt_vmzTPVfbMVmZ');
    if (success <> True) then
    begin
      MessageDlg('ErrorText' + CkCrypt2__lastErrorText(crypt), mtWarning, [mbOK], 0);
      Exit;
    end;

    CkCrypt2_putCryptAlgorithm(crypt, 'aes');
    CkCrypt2_putCipherMode(crypt, 'cbc');
    CkCrypt2_putKeyLength(crypt, 256);
    CkCrypt2_putPaddingScheme(crypt, 3);
    CkCrypt2_putEncodingMode(crypt, 'hex');

    CkCrypt2_SetEncodedIV(crypt, p_IV, 'hex');

    CkCrypt2_SetEncodedKey(crypt, PWideChar(p_KEY), 'hex');

    //if Base64 = True then
      CryptText := ADec.DecodeString(CryptText);
    decStr := CkCrypt2__decryptStringENC(crypt, PWideChar(CryptText));
  finally
    ADec.Free;
    Result := decStr;
    CkCrypt2_Dispose(crypt);
  end;
end;

procedure TfmDataSwap.downLoadClick(Sender: TObject);
var
  fileList: TStringList;
  i: Integer;
  fileName: String;
begin
  try

    mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 開始執行出貨檔案下載。');

    Self.Update;
    Application.ProcessMessages;
    try
      fileList := TStringList.Create;
      //*** 設定宅急便 Ftp 連線資訊 & 進入子目錄 ***//
      with IdFTP1 do
      begin
        if Connected then Disconnect;
        Host    := HostIP;
        Username:= ftpUser;
        Password:= ftpPW;
        Port := ftpPort;
        //Port    := 21;

        //*** 連線 ***//
        if not Connected then Connect;
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 伺服器連線成功, 準備進入目錄...');

        ChangeDir(downDir);

        //List(fileList, '1660610201_OC_' + FormatDateTime('yyyymmdd', Now) + '_*.txt', False);
        //List(fileList, '*.txt', False);
        try
          List(fileList, customerID + '_OC_' + FormatDateTime('yyyymmdd', Now) + '_*.txt', False);
        except
          on E: Exception do
          begin
            if(containstext(e.Message,'No such file or directory')) then
              mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 無檔案可供下載。');
            exit;
          end;
        end;
      end;
      mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 進入目錄成功, 開始下載檔案...');
      //*** 將資料夾下的檔案全部下載下來並回寫到資料庫 ***//
      if fileList.Count > 0 then
      begin
        //mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 總共有 ' + IntToStr(fileList.Count) + ' 個檔案處理中。');

        for i:= 0 to Pred(fileList.Count) do
          //if Trim(fileList[I]) <> '' then
          //if (Trim(fileList[I]) <> '') and (Pos('1660610201_OC_' + FormatDateTime('yyyymmdd', Now) + '_', Trim(fileList[I])) > 0) then
          if (Trim(fileList[I]) <> '') and (Pos(customerID + '_OC_' + FormatDateTime('yyyymmdd', Now) + '_', Trim(fileList[I])) > 0) then
          begin
            fileName:= ExtractFilePath(Application.ExeName) + downDir + '\' + Trim(fileList[i]) ;
            if FileExists(fileName) then
              fileName := fileName + 'd';

            //*** 下載檔案 ***//
            IdFTP1.get(Trim(fileList[i]), fileName, True);
            if IdFTP1.Connected then IdFTP1.Disconnect;
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 檔案『' + fileList[i] + '』下載完成。');

            //*** 回寫資料庫 ***//
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 開始回寫出貨... ');
            UpdateSND(fileName);
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 回寫出貨完成... ');
            //*** 連線 ***//
            if not IdFTP1.Connected then IdFTP1.Connect;

            //*** 進入子目錄 ***//
            IdFTP1.ChangeDir(downDir);

            //*** 直接刪除已下載的檔案 ***//
            IdFTP1.Delete(Trim(fileList[I]));
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 檔案『' + fileList[i] + '』已刪除。');
          end;
      end;
    finally
      if IdFTP1.Connected then IdFTP1.Disconnect;
      //if fileList.Count > 0 then
      //  mLog.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + 'Log\' +FormatDateTime('yyyymmddHHNNSS', Now) + '_OC_log.txt');
      FreeAndNil(fileList);
      Screen.Cursor:= crDefault;
      Self.Update;
      Application.ProcessMessages;
    end;
  except
    on E: Exception do
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': Error, 下載宅配通' + Zoneid + ' 庫出貨檔案時發生錯誤，錯誤原因：' + E.Message + '。');
        Exit;
      end;
  end;
end;

function TfmDataSwap.EnCrypt(CryptText: String): String;
var
  crypt: HCkCrypt2;
  success: Boolean;
  IvHex: PWideChar;
  KeyHex: PWideChar;
  encStr: PWideChar;
  decStr: PWideChar;
  AEnc: TIdEncoderMime;
begin
  crypt := CkCrypt2_Create();
  AEnc := TIdEncoderMime.Create(Nil);
  try
    success := CkCrypt2_UnlockComponent(crypt, 'MISHIHCrypt_vmzTPVfbMVmZ');
    if (success <> True) then
    begin
      MessageDlg('ErrorText' + CkCrypt2__lastErrorText(crypt), mtWarning, [mbOK], 0);
      Exit;
    end;

    CkCrypt2_putCryptAlgorithm(crypt, 'aes');
    CkCrypt2_putCipherMode(crypt, 'cbc');
    CkCrypt2_putKeyLength(crypt, 256);
    CkCrypt2_putPaddingScheme(crypt, 3);
    CkCrypt2_putEncodingMode(crypt, 'hex');

    IvHex := PWideChar(p_IV);
    CkCrypt2_SetEncodedIV(crypt, IvHex, 'hex');

    KeyHex := PWideChar(p_KEY);
    CkCrypt2_SetEncodedKey(crypt, KeyHex, 'hex');

    encStr := CkCrypt2__encryptStringENC(crypt, PWideChar(CryptText));

    //if Base64 = True then
      Result := AEnc.Encode(PWideChar(encStr))
    //else
    //  Result := encStr;
  finally
    AEnc.Free;
    CkCrypt2_Dispose(crypt);
  end;
end;

procedure TfmDataSwap.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(pickNo);
end;

procedure TfmDataSwap.FormCreate(Sender: TObject);
var
  sConnectionString1, sConnectionString2: String;
begin
  try
    //ECDB
    {$ifdef debug}
    sConnectionString1:= 'Provider Name=Oracle;Data Source=ectest1;User ID=ecoper;Password=t8ge0lfuhbnt';
    {$endif}
    //正式機
    {$ifdef release}
    //sConnectionString1:= 'Provider Name=Oracle;Data Source=ECDB2;User ID=ecstock;Password=XESPgCGVQWys';
    sConnectionString1:= 'Provider Name=Oracle;Data Source=ECDB2;User ID=ecstock;Password=!3gTHw^D2$PF5!'; //2021.09.23 換新DB
    {$endif}

    dbUniECDB.Connected:= False;
    dbUniECDB.ConnectString:= sConnectionString1;
    dbUniECDB.Connected:= True;
    dbUniECDB.Connected:= False;


    //SR3

    {$ifdef debug}
    sConnectionString2:= 'Provider Name=Oracle;Data Source=ectest1;User ID=ecoper;Password=t8ge0lfuhbnt';
    {$endif}
    //正式機
    {$ifdef release}
    sConnectionString2:= 'Provider Name=Oracle;Data Source=ECSR3;User ID=ECSTOCKUSR;Password=SIfikIezrwVbRc';
    {$endif}

    {dbUniSR3.Connected:= False;
    dbUniSR3.ConnectString:= sConnectionString2;
    dbUniSR3.Connected:= True;
    dbUniSR3.Connected:= False;}
  except
    on E: Exception do
      begin
        dbUniECDB.Connected:= False;
        dbUniSR3.Connected:= False;
        MessageBox(Self.Handle, PChar('資料庫連線錯誤，錯誤原因: ' + #13 + E.Message + #13 + '(' + FormatDateTime( 'yyyy/mm/dd  hh:mm:ss' , Now ) + ')'), PChar(Self.Caption) , MB_OK + MB_APPLMODAL + MB_ICONERROR);
        Abort;
      end;
  end;

  if (not DirectoryExists(ExtractFilePath(Application.ExeName) + upDir)) and (not ForceDirectories(ExtractFilePath(Application.ExeName) + upDir)) then
    raise Exception.Create('無法建立 ' + upDir + ' 資料夾；請手動建立之後再做匯出!! ');
  if (not DirectoryExists(ExtractFilePath(Application.ExeName) + downDir)) and (not ForceDirectories(ExtractFilePath(Application.ExeName) + downDir)) then
    raise Exception.Create('無法建立 ' + downDir + ' 資料夾；請手動建立之後再做匯出!! ');
end;

procedure TfmDataSwap.FormShow(Sender: TObject);
var
  i, j: Integer;
  iniFileName: String;
  iniFile: TIniFile;
  sSQL, dbName: String;
begin
  {$ifdef debug}
  dbName := 'ecoper';
  {$endif}
  //正式機
  {$ifdef release}
  //dbName := 'ecstockusr';
  dbName := 'ecoper';
  {$endif}

  iniFileName := ChangeFileExt(ExtractFileName(Application.ExeName),'.ini');
  iniFileName := ExtractFilePath(Application.ExeName) + iniFileName;
  if not FileExists(iniFileName) then
  begin
    mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + '：設定檔不存在，請確認...');
    Exit;
  end;
  iniFile := TIniFile.Create(iniFileName);
  try
    Zoneid := iniFile.ReadString('Set', 'Zoneid', '76');
    if (Zoneid = '76') or (Zoneid = '79') then
      customerID := '1660610201';
    if Zoneid = '77' then
      customerID := '1660610202';
    if Zoneid = '72' then
      customerID := '1660610203';  //2022.09.14 新增72海湖倉
  finally
    IniFile.Free;
  end;

  if Zoneid = '79' then
  begin
    HostIP := '20.188.123.77';
    ftpUser := '1660610201';
    ftpPW := '@PChom3';
    ftpPort := 60001;
  end
  else
  begin
    HostIP := 'wms.e-can.com.tw';
    ftpUser := '1660610201';
    ftpPW := 'pch6102WH';
    ftpPort := 21;
  end;

  sSQL := 'select addr ' +
          '  from ' + dbName + '.encryptkey ' +
          'where extraid = ''' + Zoneid + ''' ';
  {$ifdef debug}
  OpenSQL(sSQL, cdsTmp, 1, 1);
  {$endif}
  //正式機
  {$ifdef release}
  //OpenSQL(sSQL, cdsTmp, 1, 2);
  OpenSQL(sSQL, cdsTmp, 1, 1);
  {$endif}
  if cdsTmp.IsEmpty then
  begin
    mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + '：Error, ' + Zoneid + ' 庫無法取得加密金鑰...');
    Exit;
  end;
  p_KEY := cdsTmp.FieldByName('addr').AsString;

  pickNo := TStringList.Create;

  if Zoneid = '79' then
  begin
    //黑貓要取五碼郵遞區號
    if (bCatInit = False) and (FileExists(ExtractFilePath(Application.ExeName) + 'Address.mdb')) then
    begin
      try
        //ShowMessage( '黑貓要取五碼郵遞區號' );
        CatAddrII1.ConnectionString:= Format(CAT_CONNECTION_STRING, [ExtractFilePath(Application.ExeName) + 'Address.mdb']);
        bCatInit:= True;

        if ADOQ_GetZipMapping.Active then ADOQ_GetZipMapping.Close;
        ADOQ_GetZipMapping.ConnectionString:= CatAddrII1.ConnectionString;
        if not ADOQ_GetZipMapping.Active then ADOQ_GetZipMapping.Open;

        FCat5ZipCodeVersion:= CatAddrII1.GetVersion;

      except
        on E: System.SysUtils.Exception do
          begin
            MessageBox(Application.Handle, PChar(E.Message), PChar('無法取得黑貓五碼郵遞區號相關資訊!'), MB_OK + MB_APPLMODAL + MB_ICONERROR);
            Exit;
          end;
      end;
    end
    else
    begin
      MessageBox(Application.Handle, PChar('無法取得黑貓五碼郵遞區號相關資訊'), PChar('系統錯誤'), MB_OK + MB_APPLMODAL + MB_ICONERROR);
      Exit;
    end;
  end;

  if ParamCount >= 1 then
  begin
    if UpperCase(Trim(ParamStr(1))) = 'UP' then
      upLoad.Click;

    if (FormatDateTime('hh', now) >= '09') and (FormatDateTime('hh', now) <= '22') then
    begin
      if UpperCase(Trim(ParamStr(1))) = 'DOWN' then
      begin
        sSQL := 'select ''A'' from ecoper.parac ' +    //2022.11.28 增加盤點開關
                '  where patype = ''SAKINDYN'' ' +
                '    and pano = ''Y'' ';
        OpenSQL(sSQL, cdsTmp, 1, 1);
        if cdsTmp.RecordCount > 0 then
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + '：Error, ' + Zoneid + ' 庫盤點中，不可做庫存異動...')
        else
          downLoad.Click;
      end;
    end;

    if UpperCase(Trim(ParamStr(1))) = 'CHECK' then
      CheckData;

    if Pos('Error', mLog.Text) > 0 then
    begin
      if (UpperCase(Trim(ParamStr(1))) = 'UP') or (UpperCase(Trim(ParamStr(1))) = 'DOWN') then
      begin
        {j := 0;
        for i := 1 to mLog.Lines.Count do
          if Pos('Error', mLog.Lines[i]) > 0 then
          begin
            SendSMS('0930732789', '宅配通' + Zoneid + ' 庫訂單資料上傳下載有問題: ' + mLog.Lines[i]);
            j := j + 1;
            if j > 10 then break;
          end;}
        SendMail('1', mLog.Lines.Text);
      end;
      mLog.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + 'Log\' +FormatDateTime('yyyymmddHHNNSS', Now) + '_OC_log.txt');
    end;

    fmDataSwap.Close;
  end;

end;

function TfmDataSwap.Get0003ZipCode(sOddrecaddr: String): String;
var
  sZipCode: String;
begin
  //傳入地址，取新竹物流配送用的郵遞區號
  if not Assigned(ac_003) then
    ac_003:= Taddr_Compare2.Create(Application);

  Result:= '';
  sZipCode := ac_003.Search_Tno(sOddrecaddr);
  Result:= sZipCode.Trim;
end;

function TfmDataSwap.Get0044ZipCode(sOddrecaddr: String; sType: Char): String;
var
  sZipCode, sCatZipMapping: String;
begin
  // 取網家速配郵遞區號版本
  //if sZipCodeVer044.Trim.IsEmpty then
  //  sZipCodeVer044 := TGoPChomeAddress.Instance.GetVersion;

  // 傳入地址，取網家速配配送用的郵遞區號
  Result := '';
  sCatZipMapping:= TGoPChomeAddress.Instance.GetSiteName(sOddrecaddr);
  sZipCode := TGoPChomeAddress.Instance.GetCode(sOddrecaddr);
  //Result := sZipCode;  //2022.10.07 要包含區名
  //Result := sCatZipMapping + ' ' + copy(sZipCode, 1, 2) + '-' + copy(sZipCode, 3, 3) + '-' + rightstr(sZipCode, 2);
  case sType of  //2020.10.12 要拆欄位
    '1': Result := sCatZipMapping;
    '2': Result := copy(sZipCode, 1, 2) + '-' + copy(sZipCode, 3, 3) + '-' + rightstr(sZipCode, 2);
    '3': Result := sCatZipMapping + ' ' + copy(sZipCode, 1, 2) + '-' + copy(sZipCode, 3, 3) + '-' + rightstr(sZipCode, 2);
  end;

end;

function TfmDataSwap.GetCatFiveCode(sType, sOddrecaddr: String): String;
begin
//  取黑貓宅配需要的資訊
//  sType ==> 1 表示要取黑貓五碼郵遞區號，這時候 sOddrecaddr 要傳入『送件地址』
//  sType ==> 2 表示要取黑貓站所分配站名，這時候 sOddrecaddr 要傳入『黑貓五碼郵遞區號』
  try
    Result:= '';

    if not bCatInit then Exit;

    //取黑貓五碼
    if sType = '1' then
    begin
      Result:= CatAddrII1.GetCatZip(sOddrecaddr);

      if Length(Trim(Result)) < 3 then
        Result:= '';
    end
    else
    begin
      //取黑貓站所分配站名
      if Trim(sOddrecaddr) = EmptyStr then Exit;

      if (ADOQ_GetZipMapping.Active) and (ADOQ_GetZipMapping.Locate('MappingCode', sOddrecaddr, [])) then
        Result:= ADOQ_GetZipMapping.FieldByName('BASENAME').AsString;
    end;

  except
    on E: System.SysUtils.Exception do
      begin
        MessageBox(Application.Handle, PChar(E.Message), PChar('無法取得黑貓五碼郵遞區號'), MB_OK + MB_APPLMODAL + MB_ICONERROR);
        Exit;;
      end;
  end;
end;

function TfmDataSwap.GetCodPay(pNo, ordNo: String): integer;
var
  sSQL: String;
begin
  Result := 0;
  sSQL := 'select nvl(sum(case when ospickno_m = ''' + pNo + ''' then ' +
          '                         (select odmordtotal from ecoper.order_main_' + FormatDateTime('yyyy', Date) +
          '                            where odmordid = ossurid_m ' +
          '                          union ' +
          '                          select odmordtotal from ecoper.order_main_' + IntToStr(StrToInt(FormatDateTime('yyyy', Date)) - 1) +
          '                            where odmordid = ossurid_m) ' +
          '                    else 0 end), ' +
          '           0) as paymoney ' +
          '  from (select ossurid_m, ospickno_m, oszoneid_m ' +
          '          from ecoper.outstock_m1_' + GetTableYear(Date) +
          '        where ossurid_m = ''' + ordNo + ''' ' +
          '          and (nvl(trim(oscp_m), '' '') = '' '' or osstatus_m = ''SND'') ' +
          '          and ecoper.isownkindtype(osownkind_m, ''OWNKIND_PAY'') = 1 ' +
          '        group by ossurid_m, ospickno_m, oszoneid_m ' +
          '        order by to_number(nvl(trim(ecoper.getparacdata(''OutWarePay'', oszoneid_m)), ''0'')), oszoneid_m, ospickno_m) ' +
          ' where rownum = 1 ';
  OpenSQL(sSQL, cdsTmp, 1, 1);

  if cdsTmp.FieldByName('paymoney').AsInteger > 0 then
  begin
    Result := cdsTmp.FieldByName('paymoney').AsInteger;
    sSQL := 'delete ecoper.outsendlist_a7d1 where oslordid_ad1 = ''' + ordNo + ''' ';
    OpenSQL(sSQL, cdsTmp, 2, 1);

    sSQL := 'insert into ecoper.outsendlist_a7d1 (oslpickno_ad1, oslboxnum_ad1, oslordid_ad1, oslownkind_ad1, oslpaymoney_ad1, oslmark_ad1) ' +
            '  select ospickno_m, 1 as oslboxnum_ad1, ossurid_m, osownkind_m, (select odmordtotal from ecoper.order_main_' + FormatDateTime('yyyy', Date) +
            '                                                                    where odmordid = ossurid_m ' +
            '                                                                  union ' +
            '                                                                  select odmordtotal from ecoper.order_main_' + IntToStr(StrToInt(FormatDateTime('yyyy', Date)) - 1) +
            '                                                                    where odmordid = ossurid_m) oslpaymoney_ad1, ' +
            '         ''外倉貨到付款訂單'' as oslmark_ad1 ' +
            '    from ecoper.outstock_m1_' + GetTableYear(Date) +
            '  where ospickno_m = ''' + pNo + ''' ' +
            '  group by ossurid_m, ospickno_m, osownkind_m ';
    OpenSQL(sSQL, cdsTmp, 2, 1);
  end;
end;

function TfmDataSwap.GetTableYear(dDate: TDate; iIsB4: Integer): String;
var
  sMonth: String;
begin
  // iIsB4 ==> 是否取上 B4 Table 的年份( 0.N、1.Y )
  sMonth := '06';
  if FormatDateTime('mm', dDate).ToInteger() > 6 then
    sMonth := '12';
  Result := FormatDateTime('yyyy', dDate) + sMonth;

  if iIsB4 = 1 then
  begin
    if Result.Substring(4, 2).ToInteger = 12 then
      Result := (Result.ToInteger - 6).ToString()
    else
      Result := (Result.Substring(0, 4).ToInteger - 1).ToString() + '12';
  end;
end;

procedure TfmDataSwap.lbQueryClick(Sender: TObject);
var
  sSQL, sPck, dbName1, dbName2: String;
begin

  {$ifdef debug}
  dbName1 := 'ecoper';
  dbName2 := dbName1;
  {$endif}
  //正式機
  {$ifdef release}
  //dbName1 := 'ecreport';
  //dbName2 := 'ecstockusr';
  //2022.03.17 sr3頻繁異常，改直接連ECDB
  dbName1 := 'ecoper';
  dbName2 := dbName1;
  {$endif}

  //TClientDataSet(tv1_.DataController.DataSource.DataSet).Data := Null;  不需要這行，重查會有錯誤

  if Trim(edtPCKID.Text) <> '' then
  begin
    sPck := ' and pickno_d  = ''' + Trim(edtPCKID.Text) + ''' ';
  end;

  sSQL := //'select cast(rownum as varchar(10)) rec, r.* from ( ' +
          'select pickno_d, ' +
          '       pickserno_d, ' +
          '       to_char(sysdate, ''yyyymmdd'') dt, ' +
          {'       utl_i18n.unescape_reference(regexp_substr(osmark_m, ''[^、]+'', 1, 1) || '' '') recname, ' +
          '       ' + dbName1 + '.decryptphone(regexp_substr(osmark_m, ''[^、]+'', 1, 3), ''GDjVX2aSeU3yKT3u'') rectel, ' +
          '       ' + dbName1 + '.decryptphone(regexp_substr(osmark_m, ''[^、]+'', 1, 2), ''GDjVX2aSeU3yKT3u'') recmobile, ' +
          '       regexp_substr(osmark_m, ''[^、]+'', 1, 5) reczip, ' +
          '       utl_i18n.unescape_reference(regexp_substr(osmark_m, ''[^、]+'', 1, 4) || '' '') recadd, ' + }
          //2021.10.27 modi 支援派單前顧客修改收件人資訊
          '       utl_i18n.unescape_reference(o.oddreceiver || '' '') recname, ' +
          '       ' + dbName1 + '.decryptphone(o.oddrectel, ''GDjVX2aSeU3yKT3u'') rectel, ' +
          '       ' + dbName1 + '.decryptphone(o.oddrecmobile, ''GDjVX2aSeU3yKT3u'') recmobile, ' +
          '       o.oddreczip reczip, ' +
          '       utl_i18n.unescape_reference(o.oddrecaddr || '' '') recadd, ' +
          '       pickproid_d, ' +
          '       pickproname_d, ' +
          '       ossurqty_m, ' +
          '       ''n'' invoice, ' +
          '       case when ' + dbName2 + '.isownkindtype(osownkind_m, ''OWNKIND_PAY'') = ''1'' then ''102'' ' +
          '            else ''101'' end otype, ' +
          '       nvl(Trim(regexp_substr(osmark_m,''[^、]+'', 1, 7)), ''0'') codpay, ' +
          '       ''網路家庭'' sendname, ' +
          '       ''0227000898'' sendtel, ' +
          '       ''106'' sendzip, ' +
          '       ''台北市大安區敦化南路二段105號12樓'' sendadd, ' +
          '       to_char(sysdate + 1, ''yyyymmdd'') etadt, ' +
          '       ''4'' etatm, ' +
          '       '' '' email, ' +
          '       '' '' note1, ' +
          '       case when ' + dbName2 + '.isownkindtype(osownkind_m, ''OWNKIND_PAY'') = ''1'' then ''2'' ' +
          '            else ''1'' end otype2, ' +
          '       '' '' note2, ' +
          '       '' '' serno, ' +
          '       '' '' deliveryno, ' +
          '       case when osshipway_m = ''0000000002'' then ''13'' ' +
          '            when osshipway_m = ''0000000003'' then ''15'' ' +
          '            when osshipway_m = ''0000000004'' then ''11'' ' +    //宅配通
          '            when osshipway_m = ''0000000001'' then ''14'' ' +   //2021.10.21 因應速配79倉新增
          '            when osshipway_m = ''0000000044'' then ''16'' ' +   //2022.09.15 因應79倉新增網家速配
          '            else ''00'' end shipno, ' +
          '       '' '' nul, ' +
          '       ''P1'' wh_area, ' +
          '       ''0'' gift, ' +
          '       ''0'' splitreout, ' +
          '       '' '' orgno, ' +
          '       '' '' orgsno, ' +
          '       '' '' cgoods, ' +
          '       '' '' today, ' +
          '       '' '' ticket, ' +
          '       ''01'' temptype, ' +
          '       ''*'' endmark, ' +
          '       ossurid_m, ' +
          '       pickownkind_d ' +
          '  from ecoper.pick_send_d1_' + GetTableYear(Date) +
          ' left join ecoper.outstock_m1_' + GetTableYear(Date) + ' on ossurid_m = picksurordid_d and ossurno_m = picksurordno_d ' +
          ' inner join ecoper.order_detail_' + FormatDateTime('yyyy', Date) + ' o on o.oddordid = ossurid_m and o.oddordno = ossurno_m ' +
          'where pickzoneid_d = ''' + Zoneid + ''' ' +
          '  and pickstatus_d in (''NON'', ''YET'') ' +
          '  and picksurkind_d = ''OTS'' ' +
          '  and nvl(oscp_m, '' '') = '' '' ' + sPck +
          '  and nvl(osshipway_m, '' '') <> '' '' ' +
          // and nvl(o.oddreceiver, '' '') <> '' '' ' +   //2021.11.16 modi 防止收件人資訊不完全就上傳外倉
          '  and not exists (select ''A'' from ecoper.order_detail_' + FormatDateTime('yyyy', Date) + ' where oddordid = picksurordid_d and nvl(Trim(o.oddreceiver), '' '') = '' '') ' +
          '  and not exists (select ''A'' from ecoper.outstock_m1_' + GetTableYear(Date) + ' m ' +   //盤點開關
          '                   where m.ospickno_m = pickno_d ' +
          '                     and ' + dbName2 + '.ispurcsupid(m.ossupid_m) = 1 ' +
          '                     and nvl(m.oscp_m, '' '') = '' '' ' +
          '                     and exists (select ''A'' from ecoper.parac ' +
          '                                  where patype = ''SAKINDYN'' ' +
          '                                    and pano = ''Y'')) ' ;
          //'order by pickno_d, pickserno_d, ossurid_m ';

  //2022.01.07 修正跨年切資料表ERP和倉庫切的基準不一樣問題
  sSQL := 'select cast(rownum as varchar(10)) rec, r.* from ( ' + sSQL +
          ' union ' + StringReplace(sSQL, 'order_detail_' + FormatDateTime('yyyy', Date), 'order_detail_' + IntToStr(StrToInt(FormatDateTime('yyyy', Date)) - 1), [rfReplaceAll]) +
          'order by pickno_d, pickserno_d, ossurid_m ' +
          ' ) r ';
  //sSQL := StringReplace(sSQL, 'order_detail_2022', 'order_detail_2021', [rfReplaceAll]);
  //OpenSQL(sSQL, cdsQuery, 1, 2);
  //2022.03.17 sr3頻繁異常，改直接連ECDB
  OpenSQL(sSQL, cdsQuery, 1, 1);

  //StringSaveToFile(sSQL, ExtractFilePath(Application.ExeName) + '\yfyorder.sql');
end;

procedure TfmDataSwap.lbSingleClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
    UpdateSND(OpenDialog1.FileName);

end;

function TfmDataSwap.OpenSQL(SQL: UTF8String; cds: TUniQuery; act,
  flag: Integer): String;
begin

  try
    if flag = 1 then
    begin
      cds.Connection := dbUniECDB;
      try
         try
            if not dbUniECDB.Connected then dbUniECDB.Connected:= True;
         except

         end;
      finally
         dbUniECDB.Connected:= True;
      end;
    end;
    if flag = 2 then
    begin
      cds.Connection := dbUniSR3;
      try
         try
            if not dbUniSR3.Connected then dbUniSR3.Connected:= True;
         except

         end;
      finally
         dbUniSR3.Connected:= True;
      end;
    end;

    cds.Close;
    cds.SQL.Clear;
    cds.SQL.Text := SQL;

    if act = 1 then
      cds.Open;
    if act = 2 then
      cds.ExecSQL;
  except
    on eException: Exception do
    begin
      mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Open SQL Error > ' + eException.Message);
    end;
  end;

end;

procedure TfmDataSwap.SendMail(vType, MsgTxt: String);
begin
  try
    // Screen.Cursor:= crSQLWait;
    try
      // 發送郵件
      // 設置SMTP
      // IdSMTP1.Host:= 'staff.mypchome.com.tw';    //2018.10.31 modi
      //IdSMTP1.Host := 'ecmail188.global.mypchome.com.tw';
      IdSMTP1.Host := 'ecmail.rn.global.mypchome.com.tw';   //2022.09.05 郵件主機搬遷
      IdSMTP1.Port := 25;
      // IdSMTP1.Username:= 'stk_pickup';
      // IdSMTP1.Password:= 'Q&123we@';

      // IdSMTP1.AuthType:= satDefault; //satSASL
      IdSMTP1.AuthType := satNone; // 2018.10.31 modi

      IdMessage1.Clear;
      // IdMessage1.UseNowForDate:= true;
      // IdMessage1.ContentType  := 'multipart/mixed';
      // IdMessage1.Encoding     := meMIME;

      IdMessage1.From.Address := 'stk_pickup@staff.pchome.com.tw';
      IdMessage1.From.Name := 'stk_pickup';
      // IdMessage1.From.Text:= 'stk_pickup';

      if vType = '1' then
      begin
        IdMessage1.Recipients.EMailAddresses:= 'elves@staff.pchome.com.tw';
        IdMessage1.CCList.EMailAddresses:= 'stanleytu@staff.pchome.com.tw';
      end;

      /// ////////////////////////////////////////////////////
      // IdSSLIOHandlerSocket1.PassThrough:= true;

      // IdMessage1.ContentType  := 'text/html';
      // IdSMTP1.Username:= 'cynthia';
      // IdSMTP1.Password:= 'cc0425@@';
      // IdMessage1.From.Address:= 'cynthia@staff.pchome.com.tw';
      // IdMessage1.Recipients.EMailAddresses:= 'hungsue@mail.post.gov.tw';
      // IdMessage1.CCList.EMailAddresses:= 'chung640917@gmail.com;borispong@mail.post.gov.tw';
      // IdMessage1.CCList.EMailAddresses:= 'hungsue@mail.post.gov.tw;chung640917@gmail.com;borispong@mail.post.gov.tw;sengyushih@staff.pchome.com.tw;hbug@ms62.hinet.net;a0916690567@gmail.com';

      // 附件
      // sFileName:= 'E:\Source\chung64\巴斯卡三角形\Project1.rar';

      /// ////////////////////////////////////////////////////

      // E-Mail 主旨
      if vType = '1' then
        IdMessage1.Subject := '宅配通' + Zoneid + ' 庫訂單上傳下載錯誤訊息';
      // E-Mail 內容
      // IdMessage1.Body.Assign(MsgTxt);
      IdMessage1.ContentType := 'text';
      IdMessage1.CharSet := 'UTF-8';
      IdMessage1.Body.Text := MsgTxt;

      // 夾帶附檔
      // TIdAttachmentFile.Create(IdMessage1.MessageParts , MsgTxt);
      // IdMessage1.AttachmentTempDirectory
      // TIdAttachment.Create( IdMessage1.MessageParts , sFileName );

      // IdSMTP1.Connect( 1000 );
      try
        IdSMTP1.Connect;
      except
        on E: Exception do
        begin
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) +
            '：連接郵件伺服器失敗，錯誤原因：' + E.Message + '。');
          Exit;
        end;
      end;

      if IdSMTP1.Authenticate then
      begin
        try
          IdSMTP1.Send(IdMessage1);
        except
          on E: Exception do
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + '：Email寄送失敗，錯誤原因：' + E.Message + '。');
            Exit;
          end;
        end;
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + '：E-mail 已經成功寄出 !!');
      end;
      // DeleteFile( sFileName );
    finally
      if IdSMTP1.Connected then
        IdSMTP1.Disconnect;

      // Screen.Cursor:= crDefault;
    end;
  except
    on E: Exception do
    begin
      mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + '：寄送E-mail發生錯誤，錯誤原因：' + E.Message + '。');
      Abort;
    end;
  end;
end;

procedure TfmDataSwap.SendSMS(TelNo, MsgTxt: String);
var
  aResult: array of Variant;
begin
  try
    aResult := CallSP('ecoper.pkg_sms.send',VarArrayof([TelNo,
                                                        MsgTxt,
                                                        '130239',
                                                        '09',
                                                        'SYSTEM',
                                                        '',
                                                        '']),
                                            VarArrayof(['RS']),
                                            1);
  except

  end;
end;

procedure TfmDataSwap.StringSaveToFile(AString, AFileName: String);
var
  vFile:TStringList;
begin
  vFile:=TStringList.Create;
  try
    vFile.Text:=AString;
    vFile.SaveToFile(AFileName);
  finally
    FreeAndNil(vFile);
  end;
end;

function TfmDataSwap.StrToDt(strDT: String): TDateTime;
var
  s: String;
begin
  try
    if Length(strDt) = 14 then
    begin
      s := Format('%s/%s/%s %s:%s:%s', [Copy(strDt, 1, 4), Copy(strDt, 5, 2), Copy(strDt, 7, 2), Copy(strDt, 9, 2), Copy(StrDt, 11, 2), Copy(StrDt, 13, 2)]);
      Result := StrToDateTime(s);
    end
    else
    if Length(strDt) = 8 then
    begin
      s := Format('%s/%s/%s', [Copy(strDt, 1, 4), Copy(strDt, 5, 2), Copy(strDt, 7, 2)]);
      Result := StrToDateTime(s);
    end
    else
    begin
      Result := now;
    end;
  except
    Result := now;
  end;
end;

function TfmDataSwap.TextFmt(Str: string; iLen: Integer): String;
var
  tempStr: string;
  i, iLengthtmp: Integer;
  bIsBig5: Boolean;
begin
  Result := '';
  if (Trim(Str) = '') or (iLen < 2) then
    Exit;

  tempStr := '';

  // 複製的長度，必須小於等於 iLength
  iLengthtmp := 0;

  for i := 1 to length(Str) do
  begin
    // 是否為中文字
    bIsBig5 := (Ord(Str[i]) > 127);

    // 這個字是中文字，且加上這個字的長度過長
    if bIsBig5 and (iLengthtmp + 2 > iLen) then
    begin
      Result := tempStr;
      Exit;
    end;

    tempStr := tempStr + Str[i];

    // 累加字串長度
    if bIsBig5 then
      iLengthtmp := iLengthtmp + 2
    else
      iLengthtmp := iLengthtmp + 1;

  end;
  Result := tempStr;
end;

procedure TfmDataSwap.UpdatePCK;
var
  strSQL: String;
begin
    mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 開始更新資料!!');
    cdsQuery.First;
    while not cdsQuery.Eof do  //一筆一筆更新，避免把新產生的單子更新到
    begin
      if pickNo.IndexOf(cdsQuery.FieldByName('pickno_d').AsString) <> -1 then
      begin
        cdsQuery.Next;
        continue;
      end;
      try
        strSQL := 'update ecoper.outstock_m1_' + GetTableYear(Date) + ' set osstatus_m = ''PCK'', ' +
                  '                              ospickdt_m = To_Char(sysDate, ''yyyy/mm/dd''), ' +
                  '                              ospicktm_m = To_Char(sysDate, ''hh24:mi:ss''), ' +
                  '                              osmoddt_m = To_Char(sysDate, ''yyyy/mm/dd''), ' +
                  '                              osmodtm_m = To_Char(sysDate, ''hh24:mi:ss''), ' +
                  '                              osmoduser_m = ''' + UserName + '''  ' +
                  '  where ospickno_m = ''' + cdsQuery.FieldByName('pickno_d').AsString + ''' and osstatus_m = ''YET'' and Nvl(Trim(oscp_m), '' '') = '' '' ';
        OpenSQL(strSQL, cdsTmp, 2, 1);
        strSQL := 'update ecoper.pick_send_d1_' + GetTableYear(Date) + ' set pickstatus_d = ''PCK'', pickmoddt_d = To_Char(sysDate, ''yyyy/mm/dd''), pickmoder_d = ''' + UserName + '''  ' +
                '  where pickno_d = ''' + cdsQuery.FieldByName('pickno_d').AsString + ''' and pickstatus_d = ''YET'' and Nvl(Trim(pickcp_d), '' '') = '' '' ';
        OpenSQL(strSQL, cdsTmp, 2, 1);
        strSQL := 'update ecoper.outstock_time_' + GetTableYear(Date) + ' set ospickdt_m = to_char(sysDate, ''yyyy/mm/dd''), ' +
                  '                                ospicktm_m = to_char(sysDate, ''hh24:mi:ss''), ' +
                  '                                osprtdt_m  = case when nvl(trim(osprtdt_m), '' '') = '' '' then ' +
                  '                                                       to_char(sysDate, ''yyyy/mm/dd'') ' +
                  '                                                  else osprtdt_m end, ' +
                  '                                osprttm_m  = case when nvl(trim(osprttm_m), '' '') = '' '' then ' +
                  '                                                       to_char(sysDate, ''hh24:mi:ss'') ' +
                  '                                                  else osprttm_m end ' +
                  '  where ossurid_m = ''' + cdsQuery.FieldByName('ossurid_m').AsString + ''' ' +
                  '    and ospickno_m = ''' + cdsQuery.FieldByName('pickno_d').AsString + ''' ' +
                  '    and exists (select ''A'' from ecoper.pick_send_d1 ' +
                  '                  where pickno_d = ''' + cdsQuery.FieldByName('pickno_d').AsString + ''' ' +
                  '                    and pickstatus_d = ''PCK'' ' +
                  '                    and outstock_time_' + GetTableYear(Date) + '.osshipno_m = pickshipno_d) ' +
                  '                    and nvl(trim(ospickdt_m), '' '') = '' '' ';
        OpenSQL(strSQL, cdsTmp, 2, 1);
        strSQL := 'update ecoper.pick_send_d1_' + GetTableYear(Date) + ' set pickstkuserid_d = ''' + UserID + ''', ' +
                  '                               pickstkusername_d = ''' + UserName + ''', ' +
                  '                               pickstkdt_d       = to_char(sysDate, ''yyyy/mm/dd''), ' +
                  '                               pickstktm_d       = to_char(sysDate, ''hh24:mi:ss''), ' +
                  '                               PICKDT_D = case when nvl(trim(PICKDT_D), '' '') = '' '' then ' +
                  '                                                    to_char(sysDate, ''yyyy/mm/dd'') ' +
                  '                                               else PICKDT_D end, ' +
                  '                               PICKTM_D = case when nvl(trim(PICKTM_D), '' '') = '' '' then ' +
                  '                                                    to_char(sysDate, ''hh24:mi:ss'') ' +
                  '                                          else PICKTM_D end ' +
                  '  where pickno_d = ''' + cdsQuery.FieldByName('pickno_d').AsString + ''' ' +
                  '    and nvl(trim(pickstkdt_d), '' '') = '' '' ' +
                  '    and pickstatus_d = ''PCK'' ';
        OpenSQL(strSQL, cdsTmp, 2, 1);

        strSQL := 'select pickno_d from ecoper.pick_send_d1_' + GetTableYear(Date) +
                  '  where pickno_d = ''' + cdsQuery.FieldByName('pickno_d').AsString + ''' and pickstatus_d = ''PCK'' ';
        OpenSQL(strSQL, cdsTmp, 1, 1);
        if cdsTmp.RecordCount = 0 then
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error,  揀貨單號 ' + cdsQuery.FieldByName('pickno_d').AsString + ' 資料更新不成功...');
      except
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + cdsQuery.FieldByName('pickno_d').AsString + ' 資料更新發生錯誤...');
      end;
      cdsQuery.Next;
    end;
    mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 資料更新完畢!!');
end;

procedure TfmDataSwap.UpdateSND(vfileName: String);
  function CheckStatus(picNo, status: String): Boolean;
  var
    strSQL: String;
  begin
    strSQL := 'select pickno_d, pickstatus_d, pickcp_d from ecoper.pick_send_d1_' + GetTableYear(Date) +
              '  where pickno_d = ''' + picNo + ''' ' +
              'union ' +
              'select pickno_d, pickstatus_d, pickcp_d from ecoper.pick_send_d1_' + GetTableYear(Date, 1) +
              '  where pickno_d = ''' + picNo + ''' ';
    OpenSQL(strSQL, cdsTmp, 1, 1);

    with cdsTmp do
    begin
      if (FieldByName('pickstatus_d').AsString = status) and
        (Trim(FieldByName('pickcp_d').AsString) = '') then
      begin
        Result := True;
        Exit;
      end;

      if (FieldByName('pickstatus_d').AsString = 'DEL') and
        (Trim(FieldByName('pickcp_d').AsString) <> '') then
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + picNo + ' 已取消...');
        Result := False;
        Exit;
      end;

      if FieldByName('pickstatus_d').AsString = 'YET' then
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + picNo + ' 尚未派發...');
        Result := False;
        Exit;
      end;
      if FieldByName('pickstatus_d').AsString = 'PCK' then
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + picNo + ' 揀貨中...');
        Result := False;
        Exit;
      end;
      if FieldByName('pickstatus_d').AsString = 'OUT' then
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + picNo + ' 已出庫...');
        Result := False;
        Exit;
      end;
      if FieldByName('pickstatus_d').AsString = 'SND' then
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + picNo + ' 已出貨...');
        Result := False;
        Exit;
      end;

    end;
  end;
var
  fileList, picNo, shipNo, tmpList, errNo: TStringList;
  i, j, iQty: Integer;
  aResult: array of Variant;
  sSQL, sSendWay, sShipBox, fileName, BoxTmp: String;
  vDT: TDateTime;
  pNo: String;
begin

  try
    if not FileExists(vfileName) then
    begin
      mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 找不到指定的檔案處理...');
      Exit;
    end;
    fileName := vfileName;
    while Pos('\', fileName) > 0 do
      fileName := Copy(fileName, Pos('\', fileName) + 1, Length(fileName) - Pos('\', fileName));

    fileList := TStringList.Create;
    picNo := TStringList.Create;
    shipNo := TStringList.Create;
    tmpList := TStringList.Create;
    errNo := TStringList.Create;
    //計算有幾筆揀貨單
    fileList.LoadFromFile(vfileName);
    //fileList.Sort;

    //2021.12.21 add 檢查代收金額
    pNo := '';
    try
      for i := 0 to fileList.Count - 1 do
      begin
        if pNo <> Trim(Copy(fileList[i], 11, 20)) then
        begin
          sSQL := ' select nvl(sum(oslpaymoney_ad1), 0) paymoney ' +
                  '   from ecoper.outsendlist_a7d1 ' +
                  ' where oslpickno_ad1 = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ';
                  //'   group by oslpickno_ad1 '; //2022.01.10 用group 遇到無資料會沒有資料列
          OpenSQL(sSQL, cdsTmp, 1, 1);
          if cdsTmp.FieldByName('paymoney').AsInteger <> StrToInt(Trim(Copy(fileList[i], 116, 8))) then
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 代收貨款錯誤!');
            errNo.Add(Trim(Copy(fileList[i], 11, 20)));
          end;
        end
        else
        begin
          if StrToInt(Trim(Copy(fileList[i], 116, 8))) <> 0 then
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 代收貨款錯誤!');
            errNo.Add(Trim(Copy(fileList[i], 11, 20)));
          end;
        end;
        pNo := Trim(Copy(fileList[i], 11, 20));
      end;
    except
      mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 代收貨款檢查有問題!');
      Exit;
    end;

    for i := 0 to fileList.Count - 1 do
    begin
      {if StrToInt(Trim(Copy(fileList[i], 86, 6))) > 1 then  //檢查是否單箱出貨       //2022.08.01 承穎外倉出貨OC回檔併箱邏輯調整
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 非單箱出貨... ');
        Exit;
      end; }

      //檢查數量是否正確
      if tmpList.IndexOf(Trim(Copy(fileList[i], 11, 20)) + Trim(Copy(fileList[i], 46, 40))) = -1 then
      begin
        iQty := 0;
        for j := 0 to fileList.Count - 1 do
        begin
          if (Trim(Copy(fileList[i], 11, 20)) = Trim(Copy(fileList[j], 11, 20))) and
            (Trim(Copy(fileList[i], 46, 40)) = Trim(Copy(fileList[j], 46, 40))) then
            iQty := iQty + StrToInt(Trim(Copy(fileList[i], 86, 6)));
        end;
        sSQL := ' select pickno_d, sum(pickqty_d) pickqty_d ' +
                '   from ecoper.pick_send_d1_' + GetTableYear(Date) +
                ' where pickno_d = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                '   and pickproid_d = ''' + Trim(Copy(fileList[i], 46, 40)) + ''' ' +
                '   and pickstatus_d <> ''DEL'' ' +
                '   and nvl(Trim(pickcp_d), '' '') = '' '' ' +
                ' group by pickno_d ' +
                ' union ' +
                ' select pickno_d, sum(pickqty_d) pickqty_d ' +
                '   from ecoper.pick_send_d1_' + GetTableYear(Date, 1) +
                ' where pickno_d = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                '   and pickproid_d = ''' + Trim(Copy(fileList[i], 46, 40)) + ''' ' +
                '   and pickstatus_d <> ''DEL'' ' +
                '   and nvl(Trim(pickcp_d), '' '') = '' '' ' +
                ' group by pickno_d ';
        OpenSQL(sSQL, cdsTmp, 1, 1);
        //StringSaveToFile(ssql, ExtractFilePath(Application.ExeName) + '\check.txt');
        if iQty <> cdsTmp.FieldByName('pickqty_d').AsInteger then
        begin
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 出貨數量錯誤... ');
          errNo.Add(Trim(Copy(fileList[i], 11, 20)));
        end;
        tmpList.Add(Trim(Copy(fileList[i], 11, 20)) + Trim(Copy(fileList[i], 46, 40)));
      end;

      if picNo.IndexOf(Trim(Copy(fileList[i], 11, 20))) = -1 then
        picNo.Add(Trim(Copy(fileList[i], 11, 20)));

      //回寫路線碼
      if Trim(Copy(fileList[i], 106, 10)) <> '' then
      begin
       sSQL := 'update ecoper.outstock_m1_' + GetTableYear(Date) + ' set osmark_m = osmark_m || ''、'' || ''' + Trim(Copy(fileList[i], 106, 10)) + ''' ' +
              '  where ospickno_m = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
              '    and nvl(regexp_substr(osmark_m, ''[^、]+'', 1, 8), '' '') = '' '' ';
        OpenSQL(sSQL, cdsTmp, 2, 1);
      end;
    end;

    //做出庫
    for i := 0 to picNo.Count - 1 do
    begin
      if errNo.IndexOf(picNo[i]) >= 0 then continue;

      mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + picNo[i] + ' 開始出庫... ');
      if CheckStatus(picNo[i], 'PCK') then
      begin
        try
          aResult := CallSP('ecoper.sp_yfyoutwarehouse', VarArrayof([picNo[i],
                                                                  UserID,
                                                                  fComp_Name,
                                                                  '',
                                                                  Now]),
                                                      VarArrayof(['p_RESULT']),
                                                      1);
          if Trim(VarToStr(aResult[0])) <> '' then
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + picNo[i] + ' 出庫發生錯誤 => ' + Trim(VarToStr(aResult[0])));
            errNo.Add(picNo[i]);
            continue;
          end;
        except
          on eException: Exception do
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + picNo[i] + ' 呼叫 yfyoutwarehouse 發生錯誤 > ' + eException.Message);
            errNo.Add(picNo[i]);
            continue;
          end;
        end;
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + picNo[i] + ' 出庫結束... ');
      end;
      //else
      //  mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + picNo[i] + ' 狀態錯誤!');
    end;

    //做出貨
    //1.新增宅配單
    for i := 0 to fileList.Count - 1 do
    begin
      if errNo.IndexOf(Trim(Copy(fileList[i], 11, 20))) >= 0 then continue;
      if CheckStatus(Trim(Copy(fileList[i], 11, 20)), 'OUT') and (shipNo.IndexOf(Trim(Copy(fileList[i], 31, 15))) = -1) then
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' ,宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 開始新增宅配單... ');
        sSQL := ' select pickno_d, osshipway_m, case when osshipway_m = ''0000000002'' then substr(regexp_substr(osmark_m, ''[^、]+'', 1, 8), 3, 6) ' +
                '                                    when osshipway_m = ''0000000003'' then regexp_substr(osmark_m, ''[^、]+'', 1, 8) ' +
                '                                    when osshipway_m = ''0000000001'' then regexp_substr(osmark_m, ''[^、]+'', 1, 8) ' +
                '                                    when osshipway_m = ''0000000004'' then regexp_substr(osmark_m, ''[^、]+'', 1, 8) ' +  //2023.02.08 add 增加宅配通
                '                                    else regexp_substr(osmark_m, ''[^、]+'', 1, 5) end reczip, ' +
                '        case when osshipway_m <> ''0000000002'' then regexp_substr(osmark_m, ''[^、]+'', 1, 6) ' +
                '             else substr(regexp_substr(osmark_m, ''[^、]+'', 1, 8), 1, 2) end recrouting ' +
                '   from ecoper.pick_send_d1_' + GetTableYear(Date) +
                ' left join ecoper.outstock_m1_' + GetTableYear(Date) + ' on ossurid_m = picksurordid_d and ossurno_m = picksurordno_d ' +
                ' where pickno_d = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                '   and nvl(Trim(oscp_m), '' '') = '' '' ' +
                ' union ' +
                ' select pickno_d, osshipway_m, case when osshipway_m = ''0000000002'' then substr(regexp_substr(osmark_m, ''[^、]+'', 1, 8), 3, 6) ' +
                '                                    when osshipway_m = ''0000000003'' then regexp_substr(osmark_m, ''[^、]+'', 1, 8) ' +
                '                                    when osshipway_m = ''0000000001'' then regexp_substr(osmark_m, ''[^、]+'', 1, 8) ' +
                '                                    when osshipway_m = ''0000000004'' then regexp_substr(osmark_m, ''[^、]+'', 1, 8) ' +  //2023.02.08 add 增加宅配通
                '                                    else regexp_substr(osmark_m, ''[^、]+'', 1, 5) end reczip, ' +
                '        case when osshipway_m <> ''0000000002'' then regexp_substr(osmark_m, ''[^、]+'', 1, 6) ' +
                '             else substr(regexp_substr(osmark_m, ''[^、]+'', 1, 8), 1, 2) end recrouting ' +
                '   from ecoper.pick_send_d1_' + GetTableYear(Date, 1) +
                ' left join ecoper.outstock_m1_' + GetTableYear(Date, 1) + ' on ossurid_m = picksurordid_d and ossurno_m = picksurordno_d ' +
                ' where pickno_d = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                '   and nvl(Trim(oscp_m), '' '') = '' '' ';;
        OpenSQL(sSQL, cdsTmp, 1, 1);
        if Trim(cdsTmp.FieldByName('osshipway_m').AsString) = '' then
        begin
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 宅配商資料空白，無法新增宅配單!!' );
          continue;
        end;

        //vDT := StrToDT(Trim(Copy(fileList[i], 92, 14)));
        vDT := Now; //2022.06.28 因庫存對帳有問題，出貨時間改以當下系統時間
        try
          aResult := CallSP('ecoper.AddOutSendList_51', VarArrayof([cdsTmp.FieldByName('pickno_d').AsString,
                                                                    cdsTmp.FieldByName('osshipway_m').AsString,
                                                                    Trim(Copy(fileList[i], 31, 15)),
                                                                    cdsTmp.FieldByName('reczip').AsString,
                                                                    cdsTmp.FieldByName('recrouting').AsString,
                                                                    'Y',
                                                                    UserName,
                                                                    vDT,
                                                                    30,
                                                                    UserID]),
                                                        VarArrayof(['sResultCode']),
                                                        1);
          if Trim(VarToStr(aResult[0])) <> '' then
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 無法新增 ' + Trim(Copy(fileList[i], 31, 15)) + ' 宅配單資料 =>' + Trim(VarToStr(aResult[0])));
            continue;
          end;
          //2021.11.11 add 更新印宅單時間, 出庫時間
          sSQL := 'update ecoper.pick_send_d1_' + GetTableYear(Date) + ' set pickprtdt_d = ''' + FormatDateTime('yyyy/mm/dd', vDT) + ''', pickprttm_d = ''' + FormatDateTime('hh:mm:ss', vDT) + ''',  ' +
                  '                                                          pickoutdt_d = ''' + FormatDateTime('yyyy/mm/dd', vDT) + ''', pickouttm_d = ''' + FormatDateTime('hh:mm:ss', vDT) + '''  ' +
                  '  where pickno_d = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ';
          OpenSQL(sSQL, cdsTmp, 2, 1);
          //2022.1.11 add 更新印宅單時間, 虫哥那不會更新, 所以由這邊更新
          sSQL := 'update ecoper.outsendlist_' + GetTableYear(Date) + ' set oslprtdt = ''' + FormatDateTime('yyyy/mm/dd', vDT) + ''', oslprttm = ''' + FormatDateTime('hh:mm:ss', vDT) + ''',  ' +
                  '                                                         oslprtcpuname = ''' + fComp_Name + ''' ' +
                  '  where oslsendno = ''' + Trim(Copy(fileList[i], 31, 15)) + ''' ';
          OpenSQL(sSQL, cdsTmp, 2, 1);
          //2022.01.12 懶得改前面結構, 只好這邊更新
          sSQL := 'update ecoper.outstock_time_' + GetTableYear(Date) + ' set osoutdt_m = ''' + FormatDateTime('yyyy/mm/dd', vDT) + ''', osouttm_m = ''' + FormatDateTime('hh:mm:ss', vDT) + ''' ' +
                    '  where ospickno_m = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ';
          OpenSQL(sSQL, cdsTmp, 2, 1);
        except
          on eException: Exception do
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 呼叫 AddOutSendList_51 發生錯誤 > ' + eException.Message);
            continue;
          end;
        end;
        shipNo.Add(Trim(Copy(fileList[i], 31, 15)));
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' ,宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 新增宅配單結束... ');
        //2022.07.05 記Log
        try
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 開始更新記錄檔... ');
          sSQL := 'select outdt from ecoper.exwarehouse_time where pickid = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ';
          OpenSQL(sSQL, cdsTmp, 1, 1);
          if cdsTmp.RecordCount = 0 then
          begin
            sSQL := 'insert into ecoper.exwarehouse_time (pickid, outdt, keyidt, filename) ' +
                      '  values (''' + Trim(Copy(fileList[i], 11, 20)) + ''', to_date(''' + Trim(Copy(fileList[i], 92, 14)) + ''', ''yyyymmddhh24miss''), to_date(''' + FormatDateTime('yyyy/mm/dd HH:NN:SS', vDt) + ''', ''yyyy/mm/dd hh24:mi:ss''), ''' + fileName + ''') ';
            OpenSQL(sSQL, cdsTmp, 2, 1);
          end
          else if cdsTmp.FieldByName('outdt').AsDateTime < StrToDT(Trim(Copy(fileList[i], 92, 14))) then
          begin
            sSQL := 'update ecoper.exwarehouse_time set outdt = to_date(''' + Trim(Copy(fileList[i], 92, 14)) + ''', ''yyyymmddhh24miss'') ' +
                      '  where pickid = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ';
            OpenSQL(sSQL, cdsTmp, 2, 1);
          end;
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 更新記錄檔成功... ');
        except
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' 更新記錄檔失敗... ');
        end;
      end;
    end;

    //2.宅配單出貨
    for i := 0 to fileList.Count - 1 do
    begin
      if errNo.IndexOf(Trim(Copy(fileList[i], 11, 20))) >= 0 then continue;
      if CheckStatus(Trim(Copy(fileList[i], 11, 20)), 'OUT') then
      begin
        BoxTmp := '';

        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' ,宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 開始出貨... ');
        sSQL := ' select oslsendno, oslstatus ' +
                '   from ecoper.outsendlist_' + GetTableYear(Date) +
                ' where oslsendno = ''' + Trim(Copy(fileList[i], 31, 15)) + ''' ' +
                ' union '+
                ' select oslsendno, oslstatus ' +
                '   from ecoper.outsendlist_' + GetTableYear(Date, 1) +
                ' where oslsendno = ''' + Trim(Copy(fileList[i], 31, 15)) + ''' ' ;
        OpenSQL(sSQL, cdsTmp, 1, 1);
        if cdsTmp.FieldByName('oslstatus').AsString = 'SND' then  //2022.08.01 承穎外倉出貨OC回檔併箱邏輯調整
        begin
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' ,宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 已出貨... ');
          if (i > 0) and (BoxTmp <> Trim(Copy(fileList[i], 124, 40)))  then
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' ,宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 箱號不同... ');

          continue;
        end;
        BoxTmp := Trim(Copy(fileList[i], 124, 40));
        sSQL := ' select ospickno_m, osshipway_m ' +
                '   from ecoper.outstock_m1_' + GetTableYear(Date) +
                ' where ospickno_m = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                '   and nvl(Trim(oscp_m), '' '') = '' '' ' +
                ' union ' +
                ' select ospickno_m, osshipway_m ' +
                '   from ecoper.outstock_m1_' + GetTableYear(Date, 1) +
                ' where ospickno_m = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                '   and nvl(Trim(oscp_m), '' '') = '' '' ';
        OpenSQL(sSQL, cdsTmp, 1, 1);
        sSendWay := cdsTmp.FieldByName('osshipway_m').AsString;

        sShipBox := '';
        sSQL := ' select pkdlstkid_d || pkdlstkno_d as IDNO ' +
                '   from ecoper.pick_deliver_d_' + GetTableYear(Date) +
                ' where pkdloutid_d = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                //2022.02.21 modi 改要用箱子
                //'   and pkdlproid = ''' + Trim(Copy(fileList[i], 46, 40)) + ''' ' +
                '   and pkdlproid = ''' + Trim(Copy(fileList[i], 124, 40)) + ''' ' +
                '   and not exists (select ''A'' ' +
                '                     from ecoper.outsendlist_big ' +
                '                   where oslordid_b = pkdlsurordid_d ' +
                '                     and oslordno_b = pkdlsurordno_d ' +
                '                     and oslstockid_b = pkdlstkid_d ' +
                '                     and oslstockno_b = pkdlstkno_d) ' +
                '   and rownum = 1 ' +
                ' union all ' +
                ' select pkdlstkid_d || pkdlstkno_d as IDNO ' +
                '   from ecoper.pick_deliver_d_' + GetTableYear(Date, 1) +
                ' where pkdloutid_d = ''' + Trim(Copy(fileList[i], 11, 20)) + ''' ' +
                //2022.02.21 modi 改要用箱子
                //'   and pkdlproid = ''' + Trim(Copy(fileList[i], 46, 40)) + ''' ' +
                '   and pkdlproid = ''' + Trim(Copy(fileList[i], 124, 40)) + ''' ' +
                '   and not exists (select ''A'' ' +
                '                     from ecoper.outsendlist_big ' +
                '                   where oslordid_b = pkdlsurordid_d ' +
                '                     and oslordno_b = pkdlsurordno_d ' +
                '                     and oslstockid_b = pkdlstkid_d ' +
                '                     and oslstockno_b = pkdlstkno_d) ' +
                '   and rownum = 1 ';
        OpenSQL(sSQL, cdsTmp, 1, 1);
        if cdsTmp.RecordCount > 0 then
          sShipBox := Trim(cdsTmp.FieldByName('IDNO').AsString);

        //2022.02.21 modi 改要用箱子
        sSQL := ' select sbbarcode_m ' +
                '   from ecoper.shipbox_m ' +
                ' where sbbarcode_m = ''' + Trim(Copy(fileList[i], 124, 40)) + ''' ';
        OpenSQL(sSQL, cdsTmp, 1, 1);
        if cdsTmp.RecordCount > 0 then
          sShipBox := Trim(cdsTmp.FieldByName('sbbarcode_m').AsString);

        if sShipBox = '' then
        begin
          mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' ,宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 箱號無法辨識! ');
          continue;
        end;

        //vDT := StrToDT(Trim(Copy(fileList[i], 92, 14)));
        vDT := Now; //2022.06.28 因庫存對帳有問題，出貨時間改以當下系統時間
        try
          aResult := CallSP('ecoper.UpdateForShipSendII', VarArrayof([sSendWay,
                                                                      Trim(Copy(fileList[i], 31, 15)),
                                                                      fComp_Name,
                                                                      sShipBox,
                                                                      UserID,
                                                                      UserName,
                                                                      '',
                                                                      vDT]),
                                                          VarArrayof(['ResultCode']),
                                                          1);
          if Pos('Error:', Trim(VarToStr(aResult[0]))) > 0 then
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 無法出貨 >> ' + Trim(VarToStr(aResult[0])));
            continue;
          end;
        except
          on eException: Exception do
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 呼叫 UpdateForShipSendII 發生錯誤 > ' + eException.Message);
            continue;
          end;
        end;
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 揀貨單號 ' + Trim(Copy(fileList[i], 11, 20)) + ' ,宅配單號 ' + Trim(Copy(fileList[i], 31, 15)) + ' 出貨結束... ');
      end;
    end;
  finally
    FreeAndNil(fileList);
    FreeAndNil(tmpList);
    FreeAndNil(picNo);
    FreeAndNil(shipNo);
    FreeAndNil(errNo);
  end;
  //
end;

procedure TfmDataSwap.upLoadClick(Sender: TObject);
  function FillStr(str: String; len: Integer): String;
  begin
    Result := DupeString(' ', len - length(AnsiString(Trim(str)))) + Trim(str);
  end;
  function FindRec(fileName: String): Integer;
  var
    sr: TSearchRec;
    iNo: Integer;
  begin
    iNo := 0;
    if FindFirst(fileName, faAnyFile, sr) = 0 then
    begin
      repeat
        iNo := iNo + 1;
      until FindNext(sr) <> 0;
      FindClose(sr);
    end;
    Result := iNo;
  end;
  function GetHTMLDecode(str: String): String;
  begin
    //解決網頁上輸入簡體字，列印宅配單是亂碼的情況 ( 大陸的訂單 )
    try
      if Pos('&#', str) <> 0 then
        Result:= THTMLEncoding.HTML.Decode(str)
      else
        Result:= str;
    except
      Result:= str;
      Exit;
    end;
  end;
var
  tempList: TStringList;
  addList: TStringList;
  addStr, tempStr, tmpTel, tmpZip: String;
  do_fileName, dn_fileName: String;
  fileRec: Integer;
  sSQL: String;
  otype: String;
  codpay: Integer;
  note1, note2: String;
  sCatZip, sCatZipMapping, sCatCode: String;
begin

  try
    Self.Update;
    Application.ProcessMessages;
    try
      mLog.Lines.Add(FormatDateTime( 'yyyy/mm/dd hh:mm:ss' , Now ) + ': 資料準備中...' );
      lbQuery.Click;

      if cdsQuery.IsEmpty then
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': 沒有訂單資料!!');
        Exit;
      end;

      tempList := TStringList.Create;
      addList := TStringList.Create;

      //計算同一天檔案數
      //fileRec := FindRec(ExtractFilePath(Application.ExeName) + upDir + '\' + '1660610201_DO_' + FormatDateTime('yyyymmdd', Now) + '_*.txt') + 1;
      fileRec := FindRec(ExtractFilePath(Application.ExeName) + upDir + '\' + customerID + '_DO_' + FormatDateTime('yyyymmdd', Now) + '_*.txt') + 1;

      with cdsQuery  do
      begin
        First;
        while not Eof do
        begin
          sSQL := '  select count(ossurno_m) as icount ' +    //2021.08.12 add 虫哥交待，避免ERP重覆轉單
                 '         from (select ossurid_m, ossurno_m, ossurqty_m, (select oddqty from ecoper.order_detail_' + FormatDateTime('yyyy', Date) + ' where oddordid = ossurid_m and oddordno = ossurno_m ' +
                 '                                                         union ' +
                 '                                                         select oddqty from ecoper.order_detail_' + IntToStr(StrToInt(FormatDateTime('yyyy', Date)) - 1) + ' where oddordid = ossurid_m and oddordno = ossurno_m) - ' +
                 '                      (select nvl(sum(rfdqty), 0) from ecoper.refund_detail where rfdordid = ossurid_m and rfdordno = ossurno_m and rfdstatus >= 9100 and rfdstatus <> 9110) as oddqty, ' +
                 '                      ecoper.GetStockSndQty(ossurid_m, ossurno_m) as stockqty ' +
                 '                 from ecoper.outstock_m1_' + GetTableYear(Date) +
                 '                where ospickno_m = ''' + FieldByName('pickno_d').AsString + ''') ' +
                 '       where stockqty > oddqty ' ;
          OpenSQL(sSQL, cdsTmp, 1, 1);
          if cdsTmp.FieldByName('icount').AsInteger > 0 then
          begin
            mLog.Lines.Add(FormatDateTime('yyyy/mm/dd HH:NN:SS', Now) + ': Error, 揀貨單號 ' + FieldByName('pickno_d').AsString + ' 重覆轉單!!');
            pickNo.Add(FieldByName('pickno_d').AsString);
            Next;
            continue;
          end;

          note1 := FieldByName('note1').AsString;
          note2 := FieldByName('note2').AsString;
          if Zoneid = '79' then
          begin
            if FieldByName('shipno').AsString = '13' then    //黑貓
            begin
              //*** 取黑貓五碼郵遞區號 ***//
              try
                sCatZip:= GetCatFiveCode('1' , Trim(FieldByName('recadd').AsString));
                if Length(Trim(sCatZip)) < 3 then
                  sCatZip:= '';
              except
                sCatZip:= '';
              end;

              //*** 新增取黑貓站所分配站名 2010/01/25 Edit by Chung64 ***//
              if not sCatZip.IsEmpty then
              begin
                sCatZipMapping:= GetCatFiveCode('2', sCatZip);
                if sCatZipMapping.Trim.IsEmpty then
                  sCatZipMapping:= GetCatFiveCode('2', Copy(sCatZip, 1, 5)); //2020.02.25 modi 與虫哥同步程式碼

                sCatCode := sCatZipMapping + sCatZip;

                if Length(sCatCode) = 7 then
                  note1 := Copy(sCatCode, 1, 2) + '-' + Copy(sCatCode, 3, 3) + '-' + Copy(sCatCode, 6, 2)
                else if Length(sCatCode) = 8 then
                  note1 := Copy(sCatCode, 1, 2) + '-' + Copy(sCatCode, 3, 3) + '-' + Copy(sCatCode, 6, 2) + '-' + Copy(sCatCode, 8, 1)
                else
                  note1 := sCatZipMapping + '-' + sCatZip;
              end;
              note2 := FCat5ZipCodeVersion;
            end;
            if FieldByName('shipno').AsString = '15' then    //新竹
            begin
              try
                sCatZip:= Get0003ZipCode(Trim(FieldByName('recadd').AsString));
              except
                sCatZip:= '';
              end;
              note1 := sCatZip;
            end;
            //2021.11.02 add 因應速配外倉增加給大榮的宅配資訊
            if FieldByName('shipno').AsString = '14' then    //大榮
            begin
              sSQL := '  select ecoper.getsendwayzip(''0000000001'', ecoper.getzipforpickno(''' + FieldByName('pickno_d').AsString + ''')) || '' '' || ecoper.getparacdata(''OSOWNKIND_M'', ''' + FieldByName('pickownkind_d').AsString + ''' as str ' +
                      '         from dual ';
              OpenSQL(sSQL, cdsTmp, 1, 1);
              note1 := cdsTmp.FieldByName('str').AsString;
            end;
            //2022.09.15 add 因應速配外倉增加給網家速配的郵碼解析
            if FieldByName('shipno').AsString = '16' then    //網家速配
            begin
              try
                sCatZip:= Get0044ZipCode(Trim(FieldByName('recadd').AsString), '2');
                note2 :=  Get0044ZipCode(Trim(FieldByName('recadd').AsString), '1');
              except
                sCatZip:= '';
                note2 :=  '';
              end;
              note1 := sCatZip;
            end;
          end;

          codpay := 0;
          //2021.12.21 add 貨到付款
          codpay := getcodpay(FieldByName('pickno_d').AsString, FieldByName('ossurid_m').AsString);

          if codpay > 0 then otype := '102' else otype := '101';

          tempStr:= FillStr(FieldByName('pickno_d').AsString, 23)        + FillStr(FieldByName('pickserno_d').AsString, 3) +
                    FillStr(FieldByName('dt').AsString, 8)               + FillStr(' ', 60) +
                    FillStr(' ', 20)                                     + FillStr(' ', 15) +
                    FillStr(' ', 5)                                      + FillStr(' ', 100) +
                    FillStr(FieldByName('pickproid_d').AsString, 20)     + FillStr(FieldByName('pickproname_d').AsString, 100) +
                    Format('%7.7d', [FieldByName('ossurqty_m').AsInteger])+ FillStr(FieldByName('invoice').AsString, 1) +
                    //FillStr(FieldByName('otype').AsString, 3)            + Format('%8.8d', [FieldByName('codpay').AsInteger]) +
                    FillStr(otype, 3)                                    + Format('%8.8d', [codpay]) +  //2021.12.21 add 貨到付款
                    FillStr(FieldByName('sendname').AsString, 60)        + FillStr(FieldByName('sendtel').AsString, 30) +
                    FillStr(FieldByName('sendzip').AsString, 5)          + FillStr(FieldByName('sendadd').AsString, 100) +
                    FillStr(FieldByName('etadt').AsString, 8)            + FillStr(FieldByName('etatm').AsString, 2) +
                    FillStr(FieldByName('email').AsString, 40)           + FillStr(note1, 60) +
                    FillStr(FieldByName('otype2').AsString, 1)           + FillStr(note2, 64) +
                    Format('%5.5d', [fileRec])                           + FillStr(FieldByName('deliveryno').AsString, 15) +
                    FillStr(FieldByName('shipno').AsString, 14)          + FillStr(FieldByName('nul').AsString, 3) +
                    FillStr(FieldByName('wh_area').AsString, 2)          + FillStr(FieldByName('gift').AsString, 1) +
                    FillStr(FieldByName('splitreout').AsString, 1)       + FillStr(FieldByName('orgno').AsString, 23) +
                    FillStr(FieldByName('orgsno').AsString, 3)           + FillStr(FieldByName('cgoods').AsString, 1) +
                    FillStr(FieldByName('today').AsString, 1)            + FillStr(FieldByName('ticket').AsString, 1) +
                    FillStr(FieldByName('temptype').AsString, 2)         + FillStr(FieldByName('endmark').AsString, 1);
          tempList.Add(tempStr);

          {if Trim(FieldByName('rectel').AsString) = '' then
            tmpTel := FieldByName('recmobile').AsString
          else
            tmpTel := FieldByName('rectel').AsString; }    //2021.11.30 modi 抓反了
          if Trim(FieldByName('recmobile').AsString) = '' then
            tmpTel := FieldByName('rectel').AsString
          else
            tmpTel := FieldByName('recmobile').AsString;

          tmpZip := Copy(FieldByName('reczip').AsString, 1, 3);

          addStr:= FillStr(FieldByName('pickno_d').AsString, 82) + FillStr(FieldByName('recname').AsString, 163) +
                   FillStr(tmpTel, 32)                           + FillStr(tmpZip, 3)                            +
                   FillStr(TextFmt(Trim(FieldByName('recadd').AsString), 64), 64)   + FillStr('4', 50);
          addList.Add(addStr);

          Next;
        end;
      end;

        mLog.Lines.Add(FormatDateTime( 'yyyy/mm/dd hh:mm:ss' , Now ) + ': 資料準備完成，開始檔案上傳中，請稍後...');
        Self.Update;
        Application.ProcessMessages;

        //*** 取要存檔的完整檔案名稱 ***//
        //do_fileName:= ExtractFilePath(Application.ExeName) + upDir + '\' + '1660610201_DO_' + FormatDateTime('yyyymmdd', Now) + '_' + Format('%2.2d', [fileRec]) + '.txt';
        //dn_fileName:= ExtractFilePath(Application.ExeName) + upDir + '\' + '1660610201_DN_' + FormatDateTime('yyyymmdd', Now) + '_' + Format('%2.2d', [fileRec]) + '.txt';
        do_fileName:= ExtractFilePath(Application.ExeName) + upDir + '\' + customerID + '_DO_' + FormatDateTime('yyyymmdd', Now) + '_' + Format('%2.2d', [fileRec]) + '.txt';
        dn_fileName:= ExtractFilePath(Application.ExeName) + upDir + '\' + customerID + '_DN_' + FormatDateTime('yyyymmdd', Now) + '_' + Format('%2.2d', [fileRec]) + '.txt';
        //*** 把要上傳的資料存成檔案 ***//
        tempList.SaveToFile(do_fileName);
        //addList.SaveToFile(dn_fileName + 'a');
        addList.Text := EnCrypt(addList.Text);    //加密
        addList.SaveToFile(dn_fileName);
        addList.Text := DeCrypt(addList.Text);    //解密
        addList.SaveToFile(dn_fileName + 'a');
        //*** 上傳 資料檔案 ***//
        {$ifdef release}
        if UpLoadFile(do_fileName) and UpLoadFile(dn_fileName) then
        {$endif}
        begin
          //*** 共上傳幾筆資料 ***//
          mLog.Lines.Add(FormatDateTime( 'yyyy/mm/dd hh:mm:ss' , Now ) + ': 本次匯出『' + VarToStr(tempList.Count) + '』筆資料。');
          UpdatePCK;
        end
        {$ifdef release}
        else
          mLog.Lines.Add( FormatDateTime( 'yyyy/mm/dd hh:mm:ss' , Now ) + ': Error, 檔案上傳失敗。' );
        {$endif}
    finally
      FreeAndNil(tempList);
      FreeAndNil(addList);
      Screen.Cursor:= crDefault;
      Self.Update;
      Application.ProcessMessages;
    end;

    mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': 檔案上傳執行完畢。');
  except
    on E : Exception do
      begin
        mLog.Lines.Add(FormatDateTime('yyyy/mm/dd hh:mm:ss', Now) + ': Error, 檔案上傳時發生錯誤，錯誤原因：' + E.Message + '。');
        Exit;
      end;
  end;
end;

function TfmDataSwap.UpLoadFile(fileName: String): Boolean;
begin
  try
    with IdFTP1 do
    begin

      if Connected then Disconnect;  //*** 斷線 ***//
      Host    := HostIP;
      Username:= ftpUser;
      Password:= ftpPW;
      Port := ftpPort;
      //Port    := PortForCat;

      //*** 連線 ***//
      if not Connected then Connect;

      try
        //*** 進入子目錄 ***//
        //if xUpLoadDir <> '' then
        ChangeDir(upDir);

        //*** 上傳檔案 ***//
        Put(fileName, ExtractFileName(fileName));
      finally
       if Connected then Disconnect;  //*** 斷線 ***//
      end;
    end;
    Result := True;
  except
    on E : Exception do
      begin
        mLog.Lines.Add(FormatDateTime( 'yyyy/mm/dd hh:mm:ss' , Now ) + ': Error, 檔案上傳時發生錯誤，錯誤原因：' + E.Message + '。');
        Result := False;
      end;
  end;
end;

end.
