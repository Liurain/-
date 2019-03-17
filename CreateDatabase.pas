unit CreateDatabase;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, CheckData, DatabaseManage;

type
  TForm4 = class(TForm)
    Label2: TLabel;
    Label21: TLabel;
    database_beginYear_Edit: TEdit;
    Label1: TLabel;
    inout_submit_btn: TButton;
    term_ComboBox: TComboBox;
    procedure inout_submit_btnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

{$R *.dfm}



procedure TForm4.inout_submit_btnClick(Sender: TObject);
var
  y : string;
  year : integer;
  flag : boolean;
  check : TCheckData;

  term_ComboBox_index : integer;
  term : string;

  db : Database;
begin
  y := database_beginYear_Edit.Text;
  flag := true;
  check := TCheckData.Create;
  year := check.year(y, flag);
  if flag = false then
  begin
//    showMessage('请输入正确的年份，例如:2017~2018学年输入2017');
    exit;
  end;


  term_ComboBox_index := term_ComboBox.ItemIndex;
  case term_ComboBox_index of
    -1:begin
      showMessage('请选择数据库对应的学期！');
      exit;
    end;
    0:begin
      term := '上学期';
    end;
    1:begin
      term := '下学期';
    end;
  end;

  db := Database.Create;
  if db.createNewDatabase(year,term) then
  begin
    showmessage('新数据库创建成功');
    Form4.Hide;
  end else begin
    showmessage('新数据库创建失败');
  end;

end;

end.
