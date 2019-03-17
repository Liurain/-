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
//    showMessage('��������ȷ����ݣ�����:2017~2018ѧ������2017');
    exit;
  end;


  term_ComboBox_index := term_ComboBox.ItemIndex;
  case term_ComboBox_index of
    -1:begin
      showMessage('��ѡ�����ݿ��Ӧ��ѧ�ڣ�');
      exit;
    end;
    0:begin
      term := '��ѧ��';
    end;
    1:begin
      term := '��ѧ��';
    end;
  end;

  db := Database.Create;
  if db.createNewDatabase(year,term) then
  begin
    showmessage('�����ݿⴴ���ɹ�');
    Form4.Hide;
  end else begin
    showmessage('�����ݿⴴ��ʧ��');
  end;

end;

end.
