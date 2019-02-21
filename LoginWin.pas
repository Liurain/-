unit LoginWin;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, AdminWin, EnterScoreWin,
   ADOOperate, Table_parameters,CheckData, DatabaseManage,StrUtils,GlobalData,
  Table_users;

type
  TForm3 = class(TForm)
    Panel: TPanel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    login_btn: TButton;
    user_Edit: TEdit;
    psw_Edit: TEdit;
    dataPath_ComboBox: TComboBox;
    creatDatabase_btn: TButton;
    procedure FormCreate(Sender: TObject);
    procedure login_btnClick(Sender: TObject);
    procedure creatDatabase_btnClick(Sender: TObject);
  private
    tb_user : Tb_users;
    procedure getDatabaseList();
    procedure setUserPower(power : string);
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.creatDatabase_btnClick(Sender: TObject);
var
  s : string;
  y : integer;
  flag : boolean;
  check : TCheckData;
  db : Database;
begin
  if InputQuery('������ѧ���', '���������ݿ�����ʾ��ѧ��(����:2017~2018ѧ������2017)', s) then
  begin
    flag := true;
    check := TCheckData.Create;
    y := check.year(s, flag);
    if flag then
    begin
      db := Database.Create;
      if db.createNewDatabase(y) then showmessage('�����ݿⴴ���ɹ�');
      getDatabaseList();
    end else begin
      creatDatabase_btnClick(Sender);
    end;
  end;
end;

procedure TForm3.FormCreate(Sender: TObject);
begin
  panel.SetBounds(0, 0, form3.Width, form3.Height);
  getDatabaseList();

  tb_user := Tb_users.create;

end;

procedure TForm3.getDatabaseList();
var
  SearchRec: TSearchRec;
  found: integer;
  path : string;

  fileList : TStringList;
  filename : string;
  i : integer;
begin
  fileList := TStringList.Create;

  path := 'database\';       //���·����database�ļ���
  found := FindFirst(path + '*.accdb', faAnyFile, SearchRec);
  while found = 0 do
  begin
    if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') and
      (SearchRec.Attr <> faDirectory) then
    begin
      filename := SearchRec.Name;
//      filename := filename.substring(0,filename.lastIndexOf('.'));
      fileList.Add(filename);
    end;
    found := FindNext(SearchRec);
  end;

  found := FindFirst(path + '*.mdb', faAnyFile, SearchRec);
  while found = 0 do
  begin
    if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') and
      (SearchRec.Attr <> faDirectory) then
    begin
      filename := SearchRec.Name;
//      filename := filename.substring(0,filename.lastIndexOf('.'));
      fileList.Add(filename);
    end;
    found := FindNext(SearchRec);
  end;
  FindClose(SearchRec);

  fileList.Sort;
  dataPath_ComboBox.Items.Clear;
  for I := 0 to fileList.Count -1 do
  begin
    dataPath_ComboBox.Items.Add(fileList[fileList.Count-1-i])
  end;
  dataPath_ComboBox.ItemIndex := -1;
  fileList.Free;
end;

procedure TForm3.login_btnClick(Sender: TObject);
var
  user : string;
  psw : string;

  parameter : Tb_parameters;

  path : string;
  identity : integer; //�û����
  power : string;     //�û�Ȩ��
  ado : TAdoOperate;
begin
  ado := TAdoOperate.Create;

  if Comparestr(dataPath_ComboBox.Text,'') = 0 then
  begin
    showMessage('��ѡ�����ݿ�');
    dataPath_ComboBox.SetFocus;
    exit;
  end;

  path := 'database\' + dataPath_ComboBox.Text;
  ado.initConnStr(path);

  if ado.testConn() then             //�������ݿ���ͨ��
  begin
    parameter := Tb_parameters.Create;
    parameter.getpar();              //�����ݿ��л�ȡȫ�ֵĲ���

    user := user_Edit.Text;
    psw := psw_Edit.Text;
    identity := tb_user.checkLogin(user, psw, power);
    case identity of
      -1:begin   //�����֤ʧ��
        showmessage('�û������������');
      end;
      0:begin    //����Ա
        setUserPower(power);
        AdminWin.form1.Show;
        //form3.Visible := false;
      end;
      1:begin    //�ɼ�¼����
        setUserPower(power);
        EnterScoreWin.form2.Show;
        //form3.Visible := false;
      end;
    end;

  end;
end;

procedure TForm3.setUserPower(power : string);
var
  I: Integer;
begin
  for I := 1 to 9 do
  begin
    if comparestr(MidStr(power,10-i,1),'1') = 0 then
    begin
      GlobalData.power[i] := true;
    end else begin
      GlobalData.power[i] := false;
    end;
  end;
end;

end.

