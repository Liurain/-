unit LoginWin;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, AdminWin, EnterScoreWin,
   ADOOperate, Table_parameters,CheckData, DatabaseManage,StrUtils,GlobalData,
  Table_users,CreateDatabase;

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
    procedure dataPath_ComboBoxDropDown(Sender: TObject);
  private
    tb_user : Tb_users;
    procedure getDatabaseList();
    procedure setUserPower(power : string);
    function setDatabaseDate(databaseName : string):boolean;
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.creatDatabase_btnClick(Sender: TObject);
begin
  Form4.show;
end;


procedure TForm3.dataPath_ComboBoxDropDown(Sender: TObject);
begin
  getDatabaseList();
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

  pMainForm: Pointer;
  Form1: TForm1;
  Form2: TForm2;
begin
  ado := TAdoOperate.Create;

  if Comparestr(dataPath_ComboBox.Text,'') = 0 then
  begin
    showMessage('��ѡ�����ݿ�');
    dataPath_ComboBox.SetFocus;
    exit;
  end;

  if setDatabaseDate(dataPath_ComboBox.Text) = false then
  begin
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
    GlobalData.enterScoreIdentity := identity;   //��ʾ��½��Ա���
    case identity of
      -1:begin   //�����֤ʧ��
        showmessage('�û������������');
      end;
      0:begin    //����Ա
        setUserPower(power);
        Form3.Hide;                            { ���������� }
        pMainForm           := @Application.MainForm;
        Pointer(pMainForm^) := nil;
        Application.CreateForm(TForm1, Form1); { ���� Form1 ������ }
        Form1.Show;                            { ��ʾ Form1����ʱ���������� }
        Form3.Destroy;                         { ���� Form3����ʱ���򲢲����˳�������Ϊ Form1 �Ѿ������������ˡ��ر� Form2 ���˳������ˣ���Ϊ��ʱ Form2 ���������� }
        Form4.Destroy;
      end;
      1:begin    //�ɼ�¼����
        setUserPower(power);

        Form3.Hide;                            { ���������� }
        pMainForm           := @Application.MainForm;
        Pointer(pMainForm^) := nil;
        Application.CreateForm(TForm2, Form2); { ���� Form2 ������ }
        Form2.Show;                            { ��ʾ Form2����ʱ���������� }
        Form3.Destroy;
        Form4.Destroy;
      end;
    end;
  end else begin
    showmessage('���ݿ�����ʧ��');
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


function TForm3.setDatabaseDate(databaseName : string):boolean;
var
  year : integer;
  term : string;

  check : TCheckData;
  flag_year : boolean;
  flag_term : boolean;
begin
  check := TCheckData.Create;
  flag_year := true;
  flag_term := false;

  year := check.year(MidStr(databaseName,1,4), flag_year);
  term := MidStr(databaseName,11,3);
  if (comparestr(term,'��ѧ��')=0) or (comparestr(term,'��ѧ��')=0) then
  begin
    flag_term := true;
  end;


  if flag_year and flag_term then
  begin
    GlobalData.year := year;
    GlobalData.term := term;
    result := true;
  end else begin
     showMessage('���ݿ����������Ϲ淶������2017��2018ѧ����ѧ�����ݿ��ļ�Ӧ����Ϊ��2017-2018-��ѧ�ڡ�');
     result := false;
  end;


end;

end.

