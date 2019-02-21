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
  if InputQuery('输入入学年份', '请输入数据库所表示的学年(例如:2017~2018学年输入2017)', s) then
  begin
    flag := true;
    check := TCheckData.Create;
    y := check.year(s, flag);
    if flag then
    begin
      db := Database.Create;
      if db.createNewDatabase(y) then showmessage('新数据库创建成功');
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

  path := 'database\';       //相对路径，database文件夹
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
  identity : integer; //用户身份
  power : string;     //用户权限
  ado : TAdoOperate;
begin
  ado := TAdoOperate.Create;

  if Comparestr(dataPath_ComboBox.Text,'') = 0 then
  begin
    showMessage('请选择数据库');
    dataPath_ComboBox.SetFocus;
    exit;
  end;

  path := 'database\' + dataPath_ComboBox.Text;
  ado.initConnStr(path);

  if ado.testConn() then             //测试数据库连通性
  begin
    parameter := Tb_parameters.Create;
    parameter.getpar();              //从数据库中获取全局的参数

    user := user_Edit.Text;
    psw := psw_Edit.Text;
    identity := tb_user.checkLogin(user, psw, power);
    case identity of
      -1:begin   //身份验证失败
        showmessage('用户名或密码错误');
      end;
      0:begin    //管理员
        setUserPower(power);
        AdminWin.form1.Show;
        //form3.Visible := false;
      end;
      1:begin    //成绩录入者
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

