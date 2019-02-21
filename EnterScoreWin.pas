unit EnterScoreWin;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Data.DB,
  Data.Win.ADODB, Vcl.Grids, Vcl.DBGrids, Table_classes, Table_course, Table_class, Table_parameters,
  Vcl.ComCtrls, CheckData, Table_scores, Table_exam,GlobalData;

type
  TForm2 = class(TForm)
    DBGrid1: TDBGrid;
    DataSource: TDataSource;
    ADOQuery: TADOQuery;

    control_Panel: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    grade_ComboBox: TComboBox;
    classID_ComboBox: TComboBox;
    Label1: TLabel;
    orderByName_ComboBox: TComboBox;
    Label5: TLabel;
    course_ComboBox: TComboBox;
    stuID_Edit: TEdit;
    stuName_Edit: TEdit;
    Label8: TLabel;
    Label9: TLabel;

    submit_btn: TButton;
    Label6: TLabel;
    score_Edit: TEdit;
    Label7: TLabel;
    examName_ComboBox: TComboBox;
    Label4: TLabel;
    term_ComboBox: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure grade_ComboBoxChange(Sender: TObject);
    procedure courseAndCalss_ComboBoxChange(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure submit_btnClick(Sender: TObject);
    procedure term_ComboBoxChange(Sender: TObject);
    procedure examName_ComboBoxChange(Sender: TObject);
    procedure orderByName_ComboBoxChange(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure score_EditKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
  private
    check : TCheckData;
    classes : Tb_classes;
    course : Tb_course;
    scores : Tb_scores;

    dataFlag : boolean;  //������ʾ���Ʋ�����������������������ʱdataFlagΪtrue��
    procedure initForm();

    procedure setDate();
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.FormCreate(Sender: TObject);
begin
  control_Panel.Height := form2.Height;

  dataFlag := false;

  check := TCheckData.Create;
  classes := Tb_classes.Create;
  course := Tb_course.Create;
  scores := Tb_scores.create;

  initForm();
end;

procedure TForm2.FormResize(Sender: TObject);
begin
  control_Panel.SetBounds(0, 0, control_Panel.Width, Form2.Height);
  DBGrid1.SetBounds(control_Panel.Width+10,10,Form2.Width-20,Form2.Height-20);
end;

procedure TForm2.FormShow(Sender: TObject);
var
  I: Integer;
begin
  for I := 1 to 9 do
  begin
    if GlobalData.power[i] then
    begin
      case i of
        1: grade_ComboBox.Items.Add('һ�꼶');
        2: grade_ComboBox.Items.Add('���꼶');
        3: grade_ComboBox.Items.Add('���꼶');
        4: grade_ComboBox.Items.Add('���꼶');
        5: grade_ComboBox.Items.Add('���꼶');
        6: grade_ComboBox.Items.Add('���꼶');
        7: grade_ComboBox.Items.Add('���꼶');
        8: grade_ComboBox.Items.Add('���꼶');
        9: grade_ComboBox.Items.Add('���꼶');
      end;
    end;

  end;
end;

procedure TForm2.initForm();
begin
  classID_ComboBox.ItemIndex := -1;
  course_ComboBox.ItemIndex := -1;
  term_ComboBox.ItemIndex := -1;
  examName_ComboBox.ItemIndex := -1;
  stuID_Edit.Text := '';
  stuName_Edit.Text := '';
  score_Edit.Text := '';

  classID_ComboBox.Enabled := false;
  course_ComboBox.Enabled := false;
  term_ComboBox.Enabled := false;
  orderByName_ComboBox.Enabled := false;
  examName_ComboBox.Enabled := false;
  stuID_Edit.Enabled := false;
  stuName_Edit.Enabled := false;
  score_Edit.Enabled := false;
  submit_btn.Enabled := false;

  scores.courseName := '';
  scores.grade := 0;
  scores.stuID := '';
  scores.stuName := '';
  scores.score := 0;
  scores.classID := '';
  scores.exam := '';          //��������

  dataFlag := false;
end;

procedure TForm2.orderByName_ComboBoxChange(Sender: TObject);
begin
  if dataFlag then      //���¼��ɼ��������úã�ֻ�Ǹı䲿�ֲ������򲻶�����������������
  begin
    examName_ComboBoxChange(Sender);
  end;
end;

procedure TForm2.courseAndCalss_ComboBoxChange(Sender: TObject);
begin
  if dataFlag then      //���¼��ɼ��������úã�ֻ�Ǹı䲿�ֲ������򲻶�����������������
  begin
    examName_ComboBoxChange(Sender);
  end else if (course_ComboBox.ItemIndex <> -1) and (classID_ComboBox.ItemIndex <> -1) then
  begin
    scores.grade := grade_ComboBox.ItemIndex+1;
    scores.courseName := course_ComboBox.Text;
    scores.classID := classID_ComboBox.Text;
    scores.getScores(ADOQuery);

    orderByName_ComboBox.Enabled := true;
    term_ComboBox.Enabled := true;
  end;
end;

procedure TForm2.DBGrid1CellClick(Column: TColumn);
begin
  if dataFlag then
  begin
    setDate();
  end;
end;

procedure TForm2.examName_ComboBoxChange(Sender: TObject);
var
  term : string;
  orderType : string;  //����ʽ
  flag : boolean;
begin
  flag := true;
  if (course_ComboBox.ItemIndex <> -1) and (classID_ComboBox.ItemIndex <> -1) then
  begin
    scores.grade := grade_ComboBox.ItemIndex+1;
    scores.courseName := course_ComboBox.Text;
    scores.classID := classID_ComboBox.Text;
    term := check.term(term_ComboBox.ItemIndex, flag);
    scores.exam := examName_ComboBox.Text + '_' + term;
    if Comparestr('ѧ��',orderByName_ComboBox.Text)=0 then
    begin
      orderType := 'stuID';
    end else if Comparestr('����',orderByName_ComboBox.Text)=0 then
    begin
      orderType := 'stuName';
    end;
    scores.getScoresOfExam(orderType, ADOQuery);

    ADOQuery.First;       //ָ���һ����¼
    setDate();
    dataFlag := true;
    score_Edit.Enabled := true;
    submit_btn.Enabled := true;
  end;
end;

procedure TForm2.score_EditKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if 13=key then //13 �ǻس�
  begin
    submit_btnClick(Sender);
  end;
end;

procedure TForm2.setDate();
begin
  stuID_Edit.Text := ADOQuery.FieldByName('ѧ��').AsString;
  stuName_Edit.Text := ADOQuery.FieldByName('����').AsString;
  score_Edit.Text := ADOQuery.FieldByName('����').AsString;
end;

procedure TForm2.grade_ComboBoxChange(Sender: TObject);
var
  flag : boolean;
  grade : integer;
begin
  initForm();         //��ԭ����ԭ������

  flag := true;
  grade := check.Grage(grade_ComboBox.ItemIndex, flag);

  // ��ѯ�꼶Ϊgrade�����а༶
  classes.selectByGrade(ADOQuery, grade);

  //����ѯ�õ��İ༶����classID_ComboBox.Items��
  classID_ComboBox.Items.Clear;
  while not ADOQuery.eof do
  begin
    classID_ComboBox.Items.Add(ADOQuery.FieldByName('���').AsString);
    ADOQuery.Next;
  end;

  //��ѯ���꼶�����пγ�
  course.getCourseByGrade(ADOQuery, grade);
  //����ѯ�õ��İ༶����course_ComboBox.Items��
  course_ComboBox.Items.Clear;
  while not ADOQuery.eof do
  begin
    course_ComboBox.Items.Add(ADOQuery.FieldByName('�γ�����').AsString);
    ADOQuery.Next;
  end;

  classID_ComboBox.Enabled := true;
  course_ComboBox.Enabled := true;
end;

procedure TForm2.submit_btnClick(Sender: TObject);
var
  row : integer;
  flag : boolean;
begin
  flag := true;
  scores.grade := grade_ComboBox.ItemIndex+1;
  scores.courseName := course_ComboBox.Text;
  scores.score := check.Score(score_Edit.Text, flag);
  scores.stuID := check.StuID(stuID_Edit.Text, flag);

  if flag then
  begin
    row := DBGrid1.DataSource.DataSet.RecNo;
    scores.setScores(ADOQuery);
    examName_ComboBoxChange(Sender);
    DBGrid1.DataSource.DataSet.RecNo := row+1;     //��¼����һ��
    setDate();
  end;


end;

procedure TForm2.term_ComboBoxChange(Sender: TObject);
var
  grade : integer;
  term : string;
  flag : boolean;

  exam : Tb_exam;
begin
  flag := true;
  exam := Tb_exam.Create;

  grade := grade_ComboBox.ItemIndex+1;
  term  := check.term(term_ComboBox.ItemIndex, flag);
  if flag then
  begin
    exam.selectExamName(grade, term, ADOQuery);
    examName_ComboBox.Items.Clear;
    while not ADOQuery.eof do
    begin
      examName_ComboBox.Items.Add(ADOQuery.FieldByName('��������').AsString);
      ADOQuery.Next;
    end;
  end;

  if dataFlag then      //���¼��ɼ��������úã�ֻ�Ǹı䲿�ֲ������򲻶�����������������
  begin
    examName_ComboBox.ItemIndex := -1;
    stuID_Edit.Text := '';
    stuName_Edit.Text := '';
    score_Edit.Text := '';
  end;
  examName_ComboBox.Enabled := true;
end;

end.
