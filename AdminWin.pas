unit AdminWin;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.Win.ADODB, Data.DB, Vcl.Grids, ADOOperate,
  Vcl.DBGrids, Vcl.StdCtrls, Vcl.ExtCtrls,enterScoreWin,
  Table_assessment, Table_classes, Table_course, Table_class, Table_parameters, Table_users,
  Table_exam, DataTransform, fileOutput, fileInput, GlobalData,DatabaseManage,
  Vcl.ComCtrls, CheckData, Vcl.Menus, Vcl.OleServer, Word2000;

type
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    classPage: TTabSheet;
    PanelControlClass: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    class_grade_ComboBox: TComboBox;
    class_classID_ComboBox: TComboBox;
    class_studentsNum_Edit: TEdit;
    class_teacher_Edit: TEdit;
    class_operate_RG: TRadioGroup;
    class_submit_btn: TButton;
    coursePage: TTabSheet;
    PanelControlCourse: TPanel;
    Label6: TLabel;
    Label8: TLabel;
    Label1: TLabel;
    Label9: TLabel;
    course_courseName_Edit: TEdit;
    course_teacher_Edit: TEdit;
    course_grade_ComboBox: TComboBox;
    course_classID_ComboBox: TComboBox;
    course_submit_btn: TButton;
    course_operate_RG: TRadioGroup;
    studentPage: TTabSheet;
    PanelControlStudents: TPanel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    stu_sex_ComboBox: TComboBox;
    stu_stuID_Edit: TEdit;
    stu_name_Edit: TEdit;
    stu_grade_ComboBox: TComboBox;
    stu_classID_ComboBox: TComboBox;
    stu_age_ComboBox: TComboBox;
    stu_submit_btn: TButton;
    stu_operate_RG: TRadioGroup;
    parPage: TTabSheet;
    Panel1: TPanel;
    Label7: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    par_averLi_base_Edit: TEdit;
    par_dev_Edit: TEdit;
    par_greatRatio_Edit: TEdit;
    par_averWen_base_Edit: TEdit;
    par_submit_btn: TButton;
    DataSource: TDataSource;
    ADOQuery: TADOQuery;
    outputPage: TTabSheet;
    Panel2: TPanel;
    report_submit_btn: TButton;
    tab_Panel: TPanel;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    dataInput_Menu: TMenuItem;
    dataOutput_Menu: TMenuItem;
    N4: TMenuItem;
    class_Menu: TMenuItem;
    course_Menu: TMenuItem;
    stu_Menu: TMenuItem;
    exam_Menu: TMenuItem;
    par_Menu: TMenuItem;
    N10: TMenuItem;
    AdminPage: TTabSheet;
    examPage: TTabSheet;
    userPanel: TPanel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    admin_newPsw_Edit: TEdit;
    admin_checkPsw_Edit: TEdit;
    admin_oldPsw_Edit: TEdit;
    admin_userName_Edit: TEdit;
    admin_submit_btn: TButton;
    scoreInput_Menu: TMenuItem;
    adminInfo_Menu: TMenuItem;
    usersPage: TTabSheet;
    user_operate_RG: TRadioGroup;
    Label31: TLabel;
    Label32: TLabel;
    userPsw_Edit: TEdit;
    userName_Edit: TEdit;
    powerPanel: TPanel;
    userPowerPanel: TPanel;
    userPower8_CheckBox: TCheckBox;
    userPower7_CheckBox: TCheckBox;
    userPower9_CheckBox: TCheckBox;
    userPower2_CheckBox: TCheckBox;
    userPower1_CheckBox: TCheckBox;
    userPower3_CheckBox: TCheckBox;
    userPower5_CheckBox: TCheckBox;
    userPower4_CheckBox: TCheckBox;
    userPower6_CheckBox: TCheckBox;
    user_submit_btn: TButton;
    Panel4: TPanel;
    exam_operate_RG: TRadioGroup;
    Label34: TLabel;
    Label35: TLabel;
    examTime_Edit: TEdit;
    examName_Edit: TEdit;
    Panel5: TPanel;
    examGradePanel: TPanel;
    examGrade8_CheckBox: TCheckBox;
    examGrade7_CheckBox: TCheckBox;
    examGrade9_CheckBox: TCheckBox;
    examGrade2_CheckBox: TCheckBox;
    examGrade1_CheckBox: TCheckBox;
    examGrade3_CheckBox: TCheckBox;
    examGrade5_CheckBox: TCheckBox;
    examGrade4_CheckBox: TCheckBox;
    examGrade6_CheckBox: TCheckBox;
    exam_submit_btn: TButton;
    exam_grade_RG: TRadioGroup;
    user_power_RG: TRadioGroup;
    Label33: TLabel;
    inputPage: TTabSheet;
    fileOutputPanel: TPanel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label38: TLabel;
    report_type_ComboBox: TComboBox;
    report_grade_ComboBox: TComboBox;
    report_classID_ComboBox: TComboBox;
    report_courseType_ComboBox: TComboBox;
    report_exam_ComboBox: TComboBox;
    promptPanel: TPanel;
    promptLabel: TLabel;
    Panel6: TPanel;
    Label18: TLabel;
    input_type_ComboBox: TComboBox;
    inout_submit_btn: TButton;
    Label19: TLabel;
    input_filePath_Edit: TEdit;
    input_getFile_Button: TButton;
    OpenDialog1: TOpenDialog;
    databasePage: TTabSheet;
    database_Menu: TMenuItem;
    Panel7: TPanel;
    Label21: TLabel;
    database_beginYear_Edit: TEdit;
    database_submit_btn: TButton;
    N2: TMenuItem;
    Label22: TLabel;
    course_courseType_ComboBox: TComboBox;
    Label36: TLabel;
    par_averWen_top_Edit: TEdit;
    par_averLi_top_Edit: TEdit;
    Label39: TLabel;
    Label42: TLabel;
    keepBack_CheckBox: TCheckBox;
    Panel8: TPanel;
    data_course_CheckBox: TCheckBox;
    data_par_CheckBox: TCheckBox;
    data_teacher_CheckBox: TCheckBox;
    data_stu_CheckBox: TCheckBox;
    data_user_CheckBox: TCheckBox;
    Label20: TLabel;
    PageControl2: TPageControl;
    TabSheet: TTabSheet;
    DBGrid1: TDBGrid;
    TextSheet: TTabSheet;
    Memo1: TMemo;
    visual_CheckBox: TCheckBox;
    Label40: TLabel;
    input_exam_ComboBox: TComboBox;
    Label41: TLabel;
    par_grade_ComboBox: TComboBox;
    Label43: TLabel;
    par_greatWeight_Edit: TEdit;
    Label44: TLabel;
    par_inferiorWeight_Edit: TEdit;
    Label45: TLabel;
    stu_stuNum_Edit: TEdit;
    fileprint_CheckBox: TCheckBox;
    Label46: TLabel;
    report_Filetype_ComboBox: TComboBox;
    Label47: TLabel;
    report_courseName_ComboBox: TComboBox;
    report_outPutLimits_ComboBox: TComboBox;
    Label48: TLabel;
    Label49: TLabel;
    par_averWenNo_Edit: TEdit;
    Label50: TLabel;
    term_ComboBox: TComboBox;
    Label37: TLabel;
    par_averLiYes_Edit: TEdit;
    Label51: TLabel;
    Label52: TLabel;
    Label53: TLabel;
    par_averWenYes_Edit: TEdit;
    Label54: TLabel;
    par_averLiNo_Edit: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure class_submit_btnClick(Sender: TObject);

    procedure grade_ComboBoxChange(Sender: TObject);
    procedure course_classID_ComboBoxChange(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure course_submit_btnClick(Sender: TObject);
    procedure stu_submit_btnClick(Sender: TObject);
    procedure par_submit_btnClick(Sender: TObject);
    procedure stu_classID_ComboBoxChange(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure class_operate_RGClick(Sender: TObject);
    procedure course_operate_RGClick(Sender: TObject);
    procedure report_submit_btnClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure admin_submit_btnClick(Sender: TObject);
    procedure exam_operate_RGClick(Sender: TObject);
    procedure exam_grade_RGClick(Sender: TObject);
    procedure exam_submit_btnClick(Sender: TObject);
    procedure user_power_RGClick(Sender: TObject);
    procedure user_operate_RGClick(Sender: TObject);
    procedure user_submit_btnClick(Sender: TObject);
    procedure dataInput_MenuClick(Sender: TObject);
    procedure dataOutput_MenuClick(Sender: TObject);
    procedure stu_MenuClick(Sender: TObject);
    procedure class_MenuClick(Sender: TObject);
    procedure course_MenuClick(Sender: TObject);
    procedure exam_MenuClick(Sender: TObject);
    procedure scoreInput_MenuClick(Sender: TObject);
    procedure adminInfo_MenuClick(Sender: TObject);
    procedure par_MenuClick(Sender: TObject);
    procedure stu_operate_RGClick(Sender: TObject);
    procedure report_type_ComboBoxChange(Sender: TObject);
    procedure input_getFile_ButtonClick(Sender: TObject);
    procedure inout_submit_btnClick(Sender: TObject);
    procedure database_MenuClick(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure keepBack_CheckBoxClick(Sender: TObject);
    procedure database_submit_btnClick(Sender: TObject);
    procedure class_grade_ComboBoxChange(Sender: TObject);
    procedure input_type_ComboBoxChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure examName_EditClick(Sender: TObject);
    procedure report_exam_ComboBoxChange(Sender: TObject);
    procedure report_courseType_ComboBoxChange(Sender: TObject);
    procedure report_grade_ComboBoxChange(Sender: TObject);
    procedure report_Filetype_ComboBoxChange(Sender: TObject);
    procedure report_outPutLimits_ComboBoxChange(Sender: TObject);
    procedure term_ComboBoxChange(Sender: TObject);


  private
    check : TCheckData;
    classes : Tb_classes;
    course : Tb_course;
    tb_stu : Tb_class;

    exam : Tb_exam;
    user : Tb_users;
    Ado : TADOOperate;

    procedure initInputPanel();
    procedure initOutputPanel();
    procedure initClassPanel();
    procedure initCoursePanel();
    procedure initStuPanel();
    procedure initExamPanel();
    procedure initScoreInputPanel();
    procedure initAdminPanel();
    procedure initParPanel();

    procedure getInfoByGrade(var ComboBox:TComboBox;grade:integer);
    function getExamGrade():string;
    function getUserPower():string;

  public
    { Public declarations }
  end;

//var
//  Form1: TForm1;

implementation

{$R *.dfm}

{***********************************窗体设置***********************************}
procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  try
    fileOutput.wordApp.quit;
    fileOutput.ExcelApp.quit;
  except
  end;

  if ADOOperate.ADOConn.Connected  then  ADOOperate.ADOConn.Close;

//  halt;                   //退出整个程序
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  check := TCheckData.Create;
  classes := Tb_classes.Create;
  course := Tb_course.Create;
  tb_stu := Tb_class.Create;
  user := Tb_users.Create;
  exam := Tb_exam.Create;

  promptPanel.SetBounds(Width,height,0,0);
  promptPanel.Visible := false;

  AdminPage.TabVisible := false;
  databasePage.TabVisible := false;
  examPage.TabVisible := false;
  inputPage.TabVisible := false;
  usersPage.TabVisible := false;
  classPage.TabVisible := false;
  parPage.TabVisible := false;
  coursePage.TabVisible := false;
  outputPage.TabVisible := false;
  studentPage.TabVisible := false;
  TabSheet.TabVisible := false;
  TextSheet.TabVisible := false;
end;

procedure TForm1.FormResize(Sender: TObject);
begin
  pageControl1.Height := Height;

  PageControl2.SetBounds(pageControl1.Width, 0, (Width-pageControl1.Width)-5, Height-10);
  DBGrid1.SetBounds(0,0,PageControl2.Width-20,PageControl2.Height-50);
  Memo1.SetBounds(0,0,PageControl2.Width-20,PageControl2.Height-50);
end;

{
    表格显示样式设计   TDBGrid
}
procedure TForm1.DBGrid1DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
  var i :integer;
    begin
//      if gdSelected in State then Exit;
//    //定义表头的字体和背景颜色：
//        for i :=0 to (Sender as TDBGrid).Columns.Count-1 do
//        begin
//          (Sender as TDBGrid).Columns[i].Title.Font.Name :='宋体'; //字体
//          (Sender as TDBGrid).Columns[i].Title.Font.Size :=9; //字体大小
//          (Sender as TDBGrid).Columns[i].Title.Font.Color :=$000000ff; //字体颜色(红色)
//          (Sender as TDBGrid).Columns[i].Title.Color :=$0000ff00; //背景色(绿色)
//        end;
//    //隔行改变网格背景色：
//      if ADOQuery.RecNo mod 2 = 0 then
//        (Sender as TDBGrid).Canvas.Brush.Color := clInfoBk //定义背景颜色
//      else
//        (Sender as TDBGrid).Canvas.Brush.Color := RGB(191, 255, 223); //定义背景颜色
//    //定义网格线的颜色：
//    DBGrid1.DefaultDrawColumnCell(Rect,DataCol,Column,State);
//      with (Sender as TDBGrid).Canvas do //画 cell 的边框
//      begin
//        Pen.Color := $00ff0000; //定义画笔颜色(蓝色)
//        MoveTo(Rect.Left, Rect.Bottom); //画笔定位
//        LineTo(Rect.Right, Rect.Bottom); //画蓝色的横线
//        Pen.Color := $0000ff00; //定义画笔颜色(绿色)
//        MoveTo(Rect.Right, Rect.Top); //画笔定位
//        LineTo(Rect.Right, Rect.Bottom); //画绿色的竖线
//      end;
end;


{***********************************参数设置***********************************}
procedure TForm1.par_submit_btnClick(Sender: TObject);
var
  flag : boolean;
  parameter : Tb_parameters;

  averWenBase1 : real;         //文科学科评定为B级别，低于年级平均分的最大分数
  averWenTop1 :real;           //文科学科评定为B级别，高于年级平均分的最大分数
  averLiBase1 : real;          //理科学科评定为B级别，低于年级平均分的最大分数
  averLiTop1 : real;           //理科学科评定为B级别，高于年级平均分的最大分数
  averWenNo : real;
  averWenYes : real;
  averLiNo : real;
  averLiYes : real;
  dev1 : real;                 //标准差不大于该参数评为A、小于评委C
  greatRatio1 : real;          //优生率高于该百分比评为A、小于该百分比评委C
  greatWeight : real;
  inferiorWeight : real;
  grade : integer;
begin
  parameter := Tb_parameters.Create;
  flag := true;

  averWenBase1 := check.par(par_averWen_base_Edit.Text, flag);
  averWenTop1 := check.par(par_averWen_top_Edit.Text, flag);
  averLiBase1 := check.par(par_averLi_base_Edit.Text, flag);
  averLiTop1 := check.par(par_averLi_top_Edit.Text, flag);
  averWenNo := check.par(par_averWenNo_Edit.Text, flag);
  averWenYes := check.par(par_averWenYes_Edit.Text, flag);
  averLiNo := check.par(par_averLiNo_Edit.Text, flag);
  averLiYes := check.par(par_averLiYes_Edit.Text, flag);
  dev1 := check.par(par_dev_Edit.Text, flag);
  greatRatio1 := check.par(par_greatRatio_Edit.Text, flag);
  greatWeight := check.par(par_greatWeight_Edit.Text, flag);
  inferiorWeight := check.par(par_inferiorWeight_Edit.Text, flag);
  grade := check.grage(par_grade_ComboBox.ItemIndex, flag);

  if flag then
  begin
    parameter.changePar(averWenBase1, averWenTop1, averLiBase1, averLiTop1,averWenNo,averWenYes,averLiNo,averLiYes, dev1,
        greatRatio1,greatWeight,inferiorWeight,grade);
    parameter.getGradePar(ADOQuery);     //表示获取1-9年级
  end else begin
    showmessage('参数设置有误，请参照参数设置说明设置参数');
  end;
end;



{***********************************菜单设置***********************************}
procedure TForm1.dataInput_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := inputPage;     //数据导入
  initInputPanel();
end;

procedure TForm1.initInputPanel;
var
  examName : string;        //考试名称
begin
  PageControl2.ActivePage := TextSheet;

  exam.selectExam(ADOQuery);
  input_exam_ComboBox.Items.Clear;
  while not ADOQuery.eof do
  begin
    examName := ADOQuery.FieldByName('考试名称').AsString;
    input_exam_ComboBox.Items.Add(examName);
    ADOQuery.Next;
  end;

  input_exam_ComboBox.Enabled := false;
  input_exam_ComboBox.ItemIndex := -1;
end;

procedure TForm1.dataOutput_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := outPutPage;    //数据导出，生成报表
  initOutputPanel;
end;

procedure TForm1.initOutputPanel;
var
  examName : string;        //考试名称
//  term : string;            //考试所在学期

  courseList : TStringList;
  i : integer;

begin
  PageControl2.ActivePage := TextSheet;
  Memo1.Lines.Clear;
//  Memo1.Lines.Append('Word导出文件:');
//  Memo1.Lines.Append('    期末评估表           ：科目');
//  Memo1.Lines.Append('    考试评估表           ：考试名称、科目');
//  Memo1.Lines.Append('    学生成绩统计表    ：年级、班号');
//  Memo1.Lines.Append('    单科成绩表           ：考试、年级、班号、科目');
//  Memo1.Lines.Append('Excel导出文件:');

   //首先获取所有年级所有课程放到Combobox里面
  courseList := TStringList.Create;
  courseList.Sorted := true;
  courseList.Duplicates := dupIgnore;  //如有重复值则放弃
  for I := 1 to 9 do
  begin
    course.getCourseByGrade(ADOQuery, i);
    while not ADOQuery.eof do
    begin
      courseList.Add(ADOQuery.FieldByName('课程名称').AsString);
      ADOQuery.Next;
    end;
  end;
  report_courseName_ComboBox.Clear;
  report_courseName_ComboBox.Items.Assign(courseList);
  courseList.Free;

  exam.selectExam(ADOQuery);
  report_exam_ComboBox.Items.Clear;
  while not ADOQuery.eof do
  begin
    examName := ADOQuery.FieldByName('考试名称').AsString;
    report_exam_ComboBox.Items.Add(examName);
    ADOQuery.Next;
  end;

  report_Filetype_ComboBox.ItemIndex := -1;
  report_exam_ComboBox.ItemIndex := -1;
  report_courseType_ComboBox.ItemIndex := -1;
  report_grade_ComboBox.ItemIndex := -1;
  report_classID_ComboBox.ItemIndex := -1;
  report_outPutLimits_ComboBox.ItemIndex := -1;
  report_courseName_ComboBox.ItemIndex := -1;

  report_Filetype_ComboBox.Enabled := false;
  report_exam_ComboBox.Enabled := false;
  report_courseType_ComboBox.Enabled := false;
  report_grade_ComboBox.Enabled := false;
  report_classID_ComboBox.Enabled := false;
  report_outPutLimits_ComboBox.Enabled := false;
  report_courseName_ComboBox.Enabled := false;

  visual_CheckBox.Checked := false;
  fileprint_CheckBox.Checked := false;

  report_classID_ComboBox.Clear;
end;

procedure TForm1.class_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := classPage;       //班级管理
  class_operate_RG.ItemIndex := -1;
  initClassPanel;
end;

procedure TForm1.initClassPanel;
var
  i : integer;
begin
  PageControl2.ActivePage := TabSheet;
  ADOQuery.Active := false;

  class_grade_ComboBox.ItemIndex := -1;
  class_classID_ComboBox.ItemIndex := -1;
  class_studentsNum_Edit.Text := '';
  class_teacher_Edit.Text := '';

  for i:=0 to PanelControlClass.ControlCount-1 do
  begin
    if (PanelControlClass.Controls[i] is TComboBox) or
      (PanelControlClass.Controls[i] is TEdit) then
    begin
      PanelControlClass.Controls[i].Enabled := false;
    end;
  end;

  class_classID_ComboBox.Clear;

  classes.selectAll(ADOQuery);
end;

procedure TForm1.course_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := coursePage;      //课程管理
  course_operate_RG.ItemIndex := -1;
  initCoursePanel;
end;

procedure TForm1.initCoursePanel;
var
  i : integer;
  rec :array[1..2] of TVarRec;
  gradeList : TStringList;
  courseList : TStringList;
  courseField : TStringField;
  gradeField : TStringField;

  count : integer;
begin
  PageControl2.ActivePage := TabSheet;
  ADOQuery.Active := false;

  course.getAllCourse(ADOQuery);

  course_grade_ComboBox.ItemIndex := -1;
  course_classID_ComboBox.ItemIndex := -1;
  course_courseType_ComboBox.ItemIndex := -1;
  course_courseName_Edit.Text := '';
  course_teacher_Edit.Text := '';

  for i:=0 to PanelControlCourse.ControlCount-1 do
  begin
    if (PanelControlCourse.Controls[i] is TComboBox) or
      (PanelControlCourse.Controls[i] is TEdit) then
    begin
      PanelControlCourse.Controls[i].Enabled := false;
    end;
  end;

  course_classID_ComboBox.Clear;
end;

procedure TForm1.stu_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := studentPage;     //学生管理
  stu_operate_RG.ItemIndex := -1;
  initStuPanel();
end;

procedure TForm1.initStuPanel();
var
  i : integer;
begin
  PageControl2.ActivePage := TabSheet;
  ADOQuery.Active := false;

  stu_grade_ComboBox.ItemIndex := -1;
  stu_classID_ComboBox.ItemIndex := -1;
  stu_stuID_Edit.Text := '';
  stu_name_Edit.Text := '';
  stu_sex_ComboBox.ItemIndex := 0;
  stu_age_ComboBox.ItemIndex := 0;

  for i:=0 to PanelControlStudents.ControlCount-1 do
  begin
    if (PanelControlStudents.Controls[i] is TComboBox) or
      (PanelControlStudents.Controls[i] is TEdit) then
    begin
      PanelControlStudents.Controls[i].Enabled := false;
    end;
  end;

  stu_classID_ComboBox.Clear;
end;

procedure TForm1.exam_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := examPage;        //考务管理
  exam_operate_RG.ItemIndex := -1;
  initExamPanel;
end;

procedure TForm1.initExamPanel;
var
  i : integer;
begin
  PageControl2.ActivePage := TabSheet;
  ADOQuery.Active := false;

  examName_Edit.Text := '';
  examName_Edit.Font.Color :=  clblack;
//  term_ComboBox.ItemIndex := -1;
  examTime_Edit.Text := '';
  exam_grade_RG.ItemIndex := -1;

  examName_Edit.Enabled := false;
//  term_ComboBox.Enabled := false;
  examTime_Edit.Enabled := false;
  exam_grade_RG.Enabled := false;

  for i:=0 to examGradePanel.ControlCount-1 do
  begin
    examGradePanel.Controls[i].Enabled := false;
  end;
end;

procedure TForm1.scoreInput_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := usersPage;       //成绩录入人员管理
  user_operate_RG.ItemIndex := -1;
  initScoreInputPanel;
end;

procedure TForm1.initScoreInputPanel;
var
  i : integer;
begin
  PageControl2.ActivePage := TabSheet;
  ADOQuery.Active := false;

  userName_Edit.Text := '';
  userPsw_Edit.Text := '';
  user_power_RG.ItemIndex := -1;

  userName_Edit.Enabled := false;
  userPsw_Edit.Enabled := false;

  for i:=0 to userPowerPanel.ControlCount-1 do
  begin
    userPowerPanel.Controls[i].Enabled := false;
  end;
end;

procedure TForm1.adminInfo_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := adminPage;       //管理员信息管理
  initAdminPanel;
end;

procedure TForm1.initAdminPanel;
begin
  PageControl2.ActivePage := TextSheet;
end;

procedure TForm1.par_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := parPage;         //参数配置
  initParPanel;
end;

procedure TForm1.initParPanel;
var
  parameter : Tb_parameters;
begin
  PageControl2.ActivePage := TabSheet;
  ADOQuery.Active := false;

  parameter := Tb_parameters.Create;
  parameter.getGradepar(ADOQuery);

  ADOQuery.First;
  par_grade_ComboBox.ItemIndex  := ADOQuery.FieldByName('年级').AsInteger - 1;
  par_averWen_base_Edit.Text    := ADOQuery.FieldByName('文科平均分下界').AsString;
  par_averWen_top_Edit.Text     := ADOQuery.FieldByName('文科平均分上界').AsString;
  par_averLi_base_Edit.Text     := ADOQuery.FieldByName('理科平均分下界').AsString;
  par_averLi_top_Edit.Text      := ADOQuery.FieldByName('理科平均分上界').AsString;
  par_averWenNo_Edit.Text       := ADOQuery.FieldByName('文科均分否决').AsString;
  par_averWenYes_Edit.Text      := ADOQuery.FieldByName('文科均分肯定').AsString;
  par_averLiNo_Edit.Text        := ADOQuery.FieldByName('理科均分否决').AsString;
  par_averLiYes_Edit.Text       := ADOQuery.FieldByName('理科均分肯定').AsString;
  par_dev_Edit.Text             := ADOQuery.FieldByName('标准差评估差值').AsString;
  par_greatRatio_Edit.Text      := ADOQuery.FieldByName('优生率评估差值').AsString;
  par_greatWeight_Edit.Text     := ADOQuery.FieldByName('优生占比').AsString;
  par_inferiorWeight_Edit.Text  := ADOQuery.FieldByName('学困占比').AsString;
end;

procedure TForm1.database_MenuClick(Sender: TObject);
begin
  PageControl1.ActivePage := databasePage;    //数据库管理
  PageControl2.ActivePage := TextSheet;
end;



{***********************************班级管理***********************************}
procedure TForm1.class_operate_RGClick(Sender: TObject);
var
  i : integer;
begin
  initClassPanel;
  case class_operate_RG.ItemIndex of
    0:begin     //增加
      class_grade_ComboBox.Enabled := true;
      class_classID_ComboBox.Enabled := true;
      class_teacher_Edit.Enabled := true;
    end;
    1:begin     //修改
      class_teacher_Edit.Enabled := true;
    end;
    2:begin     //删除
    end;
  end;
end;

procedure TForm1.class_grade_ComboBoxChange(Sender: TObject);
var
  grade : integer;    //年级
  classID : string;
  trans : CTransForm;
begin
    trans := CTransForm.Create;
    grade := class_grade_ComboBox.ItemIndex + 1;
    if (class_operate_RG.ItemIndex = 0) then     //增加
    begin
      classes.selectByGrade(ADOQuery,grade);
      ADOQuery.Last;
      if ADOQuery.RecordCount <> 0 then
      begin
        classID := ADOQuery.FieldByName('班号').AsString;
        try
          if strtoint(classID) < 1000 then
          begin
            classID := '0'+inttostr(strtoint(classID) + 1);
          end else begin
            classID := inttostr(strtoint(classID) + 1);
          end;
          class_classID_ComboBox.Text := classID;
        except
          class_classID_ComboBox.Text := '';
        end;
      end else begin
        if trans.GradeToClassID(grade) < 10 then
        begin
          classID := '0'+inttostr(trans.GradeToClassID(grade))+'01';
        end else begin
          classID := inttostr(trans.GradeToClassID(grade))+'01';
        end;
        class_classID_ComboBox.Text := classID;
      end;
      classes.selectByGrade(ADOQuery,grade);
    end;
end;

procedure TForm1.class_submit_btnClick(Sender: TObject);
var
  flag : boolean;
  mes : string;
begin
  flag := true;

  classes.grade := check.grage(class_grade_ComboBox.ItemIndex, flag);
  classes.classID := check.classID(class_classID_ComboBox.text, flag);
  classes.stuNum := check.stuNum(class_studentsNum_Edit.text, flag);
  classes.teacher := class_teacher_Edit.text;

  if flag then
  begin
    case class_operate_RG.ItemIndex of
      -1:begin
        showmessage('请选择您想要进行的操作（增加/删除/修改）');
      end;
      0:begin   //增加
       classes.addClass();
      end;
      1:begin   //修改
        classes.changeClass();
      end;
      2:begin   //删除
        mes := '此操作将会删除该班级内学生的全部信息，是否继续？';
        if MessageBox(0, PWideChar(mes), '操作选择', MB_OKCANCEL + MB_ICONQUESTION) = ID_OK then
        begin
          classes.delClass();
        end;
      end;
    end;
    classes.selectByGrade(ADOQuery,check.grage(class_grade_ComboBox.ItemIndex, flag));
    getInfoByGrade(class_classID_ComboBox, classes.grade);
    ADOQuery.Active := true;
  end;

end;

{***********************************课程管理***********************************}
procedure TForm1.course_classID_ComboBoxChange(Sender: TObject);
var
  flag : boolean;
begin
  flag := true;
  course.grade := check.grage(course_grade_ComboBox.ItemIndex, flag);
  course.classID := check.classID(course_classID_ComboBox.Text, flag);

  if flag then
    course.getCourse(ADOQuery);
end;

procedure TForm1.course_operate_RGClick(Sender: TObject);
var
  i : integer;
begin
  initCoursePanel;
  case course_operate_RG.ItemIndex of
    0:begin      //增加
      course_grade_ComboBox.Enabled := true;
      course_courseName_Edit.Enabled := true;
      course_courseType_ComboBox.Enabled := true;
    end;
    1:begin      //修改
      course_grade_ComboBox.Enabled := true;
      course_courseName_Edit.Enabled := true;
      course_classID_ComboBox.Enabled := true;
      course_teacher_Edit.Enabled := true;
      course_courseType_ComboBox.Enabled := false;
    end;
    2:begin      //删除
      course_grade_ComboBox.Enabled := true;
//      course_courseName_Edit.Enabled := true;
    end;
  end;
end;

procedure TForm1.course_submit_btnClick(Sender: TObject);
var
  flag : boolean;
begin
  flag := true;
  course.grade := check.grage(course_grade_ComboBox.ItemIndex, flag);
  course.courseName := check.courseName(course_CourseName_Edit.text, flag);
  course.courseType := check.courseType(course_courseType_ComboBox.Text, flag);

  if flag then
  begin
    case course_operate_RG.ItemIndex of
      -1:begin
        showmessage('请选择您想要进行的操作（增加/删除/修改）');
      end;
      0:begin   //增加
       course.addCourse();
      end;
      1:begin   //修改
        course.classID := check.classID(course_classID_ComboBox.text, flag);
        course.teacher := course_teacher_Edit.text;
        course.changeCourse();
      end;
      2:begin   //删除
        course.delCourse();
      end;
    end;
    course.getCourseByGrade(ADOQuery, course.grade);
  end;

end;

{***********************************学生管理***********************************}
procedure TForm1.stu_classID_ComboBoxChange(Sender: TObject);
var
  flag : boolean;
begin
  flag := true;
  tb_stu.classID := check.classID(stu_classID_ComboBox.Text,flag);
  if flag then
  begin
    stu_stuNum_Edit.Text := inttostr(classes.getStuNum(tb_stu.classID));
    tb_stu.getStu(ADOQuery);
    if stu_operate_RG.ItemIndex = 0 then
    begin
      stu_stuID_Edit.SetFocus;
    end;
  end;
end;

procedure TForm1.stu_operate_RGClick(Sender: TObject);
var
  i : integer;
begin
  initStuPanel;
  case stu_operate_RG.ItemIndex of
    0:begin          //增加
      for i:=0 to PanelControlStudents.ControlCount-1 do
      begin
        PanelControlStudents.Controls[i].Enabled := true;
      end;
    end;
    1:begin          //修改
      for i:=0 to PanelControlStudents.ControlCount-1 do
      begin
        PanelControlStudents.Controls[i].Enabled := true;
      end;
      stu_stuID_Edit.Enabled := false;
    end;
    2:begin          //删除
      stu_grade_ComboBox.Enabled := true;
      stu_classID_ComboBox.Enabled := true;
    end;
  end;
end;

procedure TForm1.stu_submit_btnClick(Sender: TObject);
var
  flag : boolean;
begin
  flag := true;

  tb_stu.classID := check.classID(stu_classID_ComboBox.Text, flag);
  tb_stu.stuID := check.stuID(stu_stuID_Edit.Text, flag);
  tb_stu.stuName := check.stuName(stu_name_Edit.Text, flag);
  tb_stu.sex := check.stuSex(stu_sex_ComboBox.ItemIndex, flag);
  tb_stu.age := check.stuAge(stu_age_ComboBox.Text, flag);

  if flag then
  begin
    case stu_operate_RG.ItemIndex of
      -1:begin
        showmessage('请选择您想要进行的操作（增加/删除/修改）');
      end;
      0:begin   //增加
        tb_stu.addStu();
      end;
      1:begin   //修改
        tb_stu.changeStu();
      end;
      2:begin   //删除
        tb_stu.delStu();
      end;
    end;
    tb_stu.getStu(ADOQuery);
  end;
end;


{********************************管理员信息修改********************************}
procedure TForm1.admin_submit_btnClick(Sender: TObject);
var
  userName : string;
  oldPsw : string;
  newPsw : string;
  checkPsw : string;
  res : integer;  //修改结果
begin
  userName := admin_userName_Edit.Text;
  oldPsw := admin_oldPsw_Edit.Text;
  newPsw := admin_newPsw_Edit.Text;
  checkPsw := admin_checkPsw_Edit.Text;

  if Comparestr(newPsw,checkPsw) = 0 then
  begin
    res := user.changeAdminInfo(userName, oldPsw, newPsw);
    case res of
      0: showMessage('旧密码不正确');
      1: showMessage('修改成功');
    end;
  end else begin
    admin_newPsw_Edit.Text := '';
    admin_checkPsw_Edit.Text := '';
    showMessage('两次输入新密码不一致');
    admin_newPsw_Edit.SetFocus;
  end;
end;


{***********************************考务管理***********************************}
procedure TForm1.exam_operate_RGClick(Sender: TObject);
begin
  initExamPanel;
  exam.selectExam(ADOQuery);
  case exam_operate_RG.ItemIndex of
    0:begin      //增加
      examName_Edit.Enabled := true;
//      term_ComboBox.Enabled := true;
      examTime_Edit.Enabled := true;
      exam_grade_RG.Enabled := true;
      examTime_Edit.Text := '';

      examName_Edit.Text := '请填写尽量简短的汉字';
      examName_Edit.Font.Color :=  clSilver;
    end;
    1:begin      //修改
      examTime_Edit.Enabled := true;
      exam_grade_RG.Enabled := true;
    end;
    2:begin      //删除
    end;
  end;
end;

procedure TForm1.examName_EditClick(Sender: TObject);
begin
  examName_Edit.Text := '';
  examName_Edit.Font.Color := clblack;
end;

procedure TForm1.exam_grade_RGClick(Sender: TObject);
var
  i : integer;
begin
  if exam_grade_RG.ItemIndex = 4 then    //自定义
  begin
    for i:=0 to examGradePanel.ControlCount-1 do
    begin
      examGradePanel.Controls[i].Enabled := true;
    end;
  end else begin
    for i:=0 to examGradePanel.ControlCount-1 do
    begin
      examGradePanel.Controls[i].Enabled := false;
    end;
  end;
end;

function TForm1.getExamGrade;
var
  buf: array[0..8] of Char;
begin
  if examGrade1_CheckBox.Checked then
  begin
    buf[8] := '1';
  end else begin
    buf[8] := '0';
  end;

  if examGrade2_CheckBox.Checked then
  begin
    buf[7] := '1';
  end else begin
    buf[7] := '0';
  end;

  if examGrade3_CheckBox.Checked then
  begin
    buf[6] := '1';
  end else begin
    buf[6] := '0';
  end;
  if examGrade4_CheckBox.Checked then
  begin
    buf[5] := '1';
  end else begin
    buf[5] := '0';
  end;
  if examGrade5_CheckBox.Checked then
  begin
    buf[4] := '1';
  end else begin
    buf[4] := '0';
  end;
  if examGrade6_CheckBox.Checked then
  begin
    buf[3] := '1';
  end else begin
    buf[3] := '0';
  end;
  if examGrade7_CheckBox.Checked then
  begin
    buf[2] := '1';
  end else begin
    buf[2] := '0';
  end;
  if examGrade8_CheckBox.Checked then
  begin
    buf[1] := '1';
  end else begin
    buf[1] := '0';
  end;
  if examGrade9_CheckBox.Checked then
  begin
    buf[0] := '1';
  end else begin
    buf[0] := '0';
  end;

  result := string(buf);
end;

procedure TForm1.exam_submit_btnClick(Sender: TObject);
var
  examGrade : string;
begin
  examGrade := '000000000';
  case exam_grade_RG.ItemIndex of
    0:begin     //所有年级
      examGrade := '111111111';
    end;
    1:begin     //1~3年级
      examGrade := '000000111';
    end;
    2:begin     //4~6年级
      examGrade := '000111000';
    end;
    3:begin     //7~9年级
      examGrade := '111000000';
    end;
    4:begin     //自定义
      examGrade := getExamGrade;
    end;
  end;

  exam.examName := examName_Edit.Text;
//  exam.term := term_ComboBox.Text;
  exam.examDate := examTime_Edit.Text;
  exam.examGrade := examGrade;
  case exam_operate_RG.ItemIndex of
    0:begin      //增加
      exam.addExam();
    end;
    1:begin      //修改
      exam.changeExam();
    end;
    2:begin      //删除
      exam.delExam(true);
    end;
  end;
  exam.selectExam(ADOQuery);
end;

{***********************************用户管理***********************************}
procedure TForm1.user_operate_RGClick(Sender: TObject);
begin
  user.selectUser(ADOQuery);
  case user_operate_RG.ItemIndex of
    0:begin      //增加
      userName_Edit.Enabled := true;
      userPsw_Edit.Enabled := true;
      userName_Edit.Text := '';
      userPsw_Edit.Text := '';
    end;
    1:begin      //修改
      userName_Edit.Enabled := false;
      userPsw_Edit.Enabled := true;
    end;
    2:begin      //删除
      userName_Edit.Enabled := false;
      userPsw_Edit.Enabled := false;
    end;
  end;
end;

procedure TForm1.user_power_RGClick(Sender: TObject);
var
  i : integer;
begin
  if user_power_RG.ItemIndex = 4 then    //自定义
  begin
    for i:=0 to userPowerPanel.ControlCount-1 do
    begin
      userPowerPanel.Controls[i].Enabled := true;
    end;
  end else begin
    for i:=0 to userPowerPanel.ControlCount-1 do
    begin
      userPowerPanel.Controls[i].Enabled := false;
    end;
  end;
end;

function TForm1.getUserPower;
var
  buf: array[0..8] of Char;
begin
  if userPower1_CheckBox.Checked then
  begin
    buf[8] := '1';
  end else begin
    buf[8] := '0';
  end;

  if userPower2_CheckBox.Checked then
  begin
    buf[7] := '1';
  end else begin
    buf[7] := '0';
  end;

  if userPower3_CheckBox.Checked then
  begin
    buf[6] := '1';
  end else begin
    buf[6] := '0';
  end;
  if userPower4_CheckBox.Checked then
  begin
    buf[5] := '1';
  end else begin
    buf[5] := '0';
  end;
  if userPower5_CheckBox.Checked then
  begin
    buf[4] := '1';
  end else begin
    buf[4] := '0';
  end;
  if userPower6_CheckBox.Checked then
  begin
    buf[3] := '1';
  end else begin
    buf[3] := '0';
  end;
  if userPower7_CheckBox.Checked then
  begin
    buf[2] := '1';
  end else begin
    buf[2] := '0';
  end;
  if userPower8_CheckBox.Checked then
  begin
    buf[1] := '1';
  end else begin
    buf[1] := '0';
  end;
  if userPower9_CheckBox.Checked then
  begin
    buf[0] := '1';
  end else begin
    buf[0] := '0';
  end;

  result := string(buf);
end;

procedure TForm1.user_submit_btnClick(Sender: TObject);
var
  userPower : string;
begin
  userPower := '000000000';
  case user_power_RG.ItemIndex of
    0:begin     //所有年级
      userPower := '111111111';
    end;
    1:begin     //1~3年级
      userPower := '000000111';
    end;
    2:begin     //4~6年级
      userPower := '000111000';
    end;
    3:begin     //7~9年级
      userPower := '111000000';
    end;
    4:begin     //自定义
      userPower := getUserPower;
    end;
  end;

  user.userName := userName_Edit.Text;
  user.password := userPsw_Edit.Text;
  user.power := userPower;
  case user_operate_RG.ItemIndex of
    0:begin      //增加
      user.addUser(ADOQuery);
    end;
    1:begin      //修改
      user.changeUser(ADOQuery);
    end;
    2:begin      //删除
      user.delUser(ADOQuery);
    end;
  end;
  user.selectUser(ADOQuery);
end;


{***********************************文件导出***********************************}
procedure TForm1.report_courseType_ComboBoxChange(Sender: TObject);
var
  allCourse : boolean;
begin
//判断是全科还是单科
  if comparestr('全科',report_courseType_ComboBox.Text) = 0 then
  begin
    allCourse := true;
  end else if comparestr('单科',report_courseType_ComboBox.Text) = 0 then
  begin
    allCourse := false;
  end else begin
    showmessage('科目类型错误');
  end;

  if allCourse then       //全科
  begin
    report_courseName_ComboBox.ItemIndex := -1;
    report_courseName_ComboBox.Enabled := false;
    if (comparestr('登分表',report_type_ComboBox.Text) = 0) then
    begin
      if (comparestr('Word',report_Filetype_ComboBox.Text) = 0) then
      begin
        showmessage('全科登分表不支持输出word文件，文件类型自动调整为Excel。');
      end;
      report_Filetype_ComboBox.Enabled := false;
      report_Filetype_ComboBox.ItemIndex := 1;      //1表示Excel
    end;

    if (comparestr('成绩表',report_type_ComboBox.Text) = 0) then
    begin
      if (comparestr('Word',report_Filetype_ComboBox.Text) = 0) then
      begin
        showmessage('全科成绩表不支持输出word文件，文件类型自动调整为Excel。');
      end;
      report_Filetype_ComboBox.Enabled := false;
      report_Filetype_ComboBox.ItemIndex := 1;      //1表示Excel
    end;
  end else begin         //单科
    if (comparestr('登分表',report_type_ComboBox.Text) <> 0) then
    begin
      report_courseName_ComboBox.Enabled := true;
    end;
    report_Filetype_ComboBox.Enabled := true;
  end;
end;

procedure TForm1.report_exam_ComboBoxChange(Sender: TObject);
begin
//  report_courseName_ComboBox.SetFocus;    //report_courseName_ComboBox中有所有年级课程
end;

procedure TForm1.report_Filetype_ComboBoxChange(Sender: TObject);
begin
  if  report_type_ComboBox.ItemIndex <> -1 then      //输出文件种类选好了
  begin
    report_courseType_ComboBox.Enabled := true;
  end;
end;

procedure TForm1.report_grade_ComboBoxChange(Sender: TObject);
var
  ADOQuery1 : TADOQuery;
  grade : integer;
begin
  ADOQuery1 := TADOQuery.Create(nil);
  grade := report_grade_ComboBox.ItemIndex+1;
//填充班号
  report_classID_ComboBox.Items.Clear;
  classes.selectByGrade(ADOQuery1, grade);
  while not ADOQuery1.eof do
  begin
    report_classID_ComboBox.Items.Add(ADOQuery1.FieldByName('班号').AsString);
    ADOQuery1.Next;
  end;
//更新课程名称
  report_courseName_ComboBox.Clear;
  course.getCourseByGrade(ADOQuery1, grade);
  while not ADOQuery1.eof do
  begin
    report_courseName_ComboBox.Items.Add(ADOQuery1.FieldByName('课程名称').AsString);
    ADOQuery1.Next;
  end;
end;

procedure TForm1.report_outPutLimits_ComboBoxChange(Sender: TObject);
var
  outFileFlag : integer;        //0全校、1全年级、2某一班级
begin
  outFileFlag := report_outPutLimits_ComboBox.ItemIndex;
  case outFileFlag of
    0: begin
      report_grade_ComboBox.ItemIndex := -1;
      report_grade_ComboBox.Enabled := false;
      report_classID_ComboBox.ItemIndex := -1;
      report_classID_ComboBox.Enabled := false;
    end;
    1: begin
      report_grade_ComboBox.Enabled := true;
      report_classID_ComboBox.ItemIndex := -1;
      report_classID_ComboBox.Enabled := false;
    end;
    2: begin
      report_grade_ComboBox.Enabled := true;
      report_classID_ComboBox.Enabled := true;
    end;
  end;

end;

procedure TForm1.report_submit_btnClick(Sender: TObject);
var
  output1 : output;
  assessment : Tb_assessment;

  exam : string;
  course : string;
  grade : integer;
  classID : string;

  i : integer;
begin
//检查参数配置的完整性
  for i:=0 to fileOutputPanel.ControlCount-1 do
  begin
    if (fileOutputPanel.Controls[i] is TComboBox) and
      (fileOutputPanel.Controls[i].Enabled) then
    begin
      if TComboBox(fileOutputPanel.Controls[i]).ItemIndex = -1 then
      begin
        showMessage('数据配置不完整，请检查参数配置。');
        exit;
      end;
    end;
  end;

//显示提示信息，准备导出文件
  promptPanel.SetBounds(round((Width-promptLabel.Width)/2),
      round((Height-promptLabel.Height)/2),promptLabel.Width,promptLabel.Height);
  promptLabel.SetBounds(0,0,promptPanel.Width,promptPanel.Height);
  promptLabel.Caption := '数据导出中，请稍后...';
  promptPanel.Visible := true;
  promptPanel.BringToFront;

//根据参数导出文件
  output1 := output.create;
  output1.isVisual(visual_CheckBox.Checked);
  output1.isfilePrint(fileprint_CheckBox.Checked);
  assessment := Tb_assessment.Create;

  case report_type_ComboBox.ItemIndex of
    0: begin    //登分表
      case report_Filetype_ComboBox.ItemIndex of
        0:begin                                                     //word
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                 //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {登分表、word、全科、全校}
                end;
                1:begin                                             //全年级
                  {登分表、word、全科、全年级}
                end;
                2:begin                                             //班级
                  {登分表、word、全科、班级}
                end;
              end;
            end;
            1:begin                                                 //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {登分表、word、单科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  output1.stat_word_aCourse_school();
                end;
                1:begin                                             //全年级
                  {登分表、word、单科、全年级}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  output1.stat_word_aCourse_grade(report_grade_ComboBox.ItemIndex + 1);
                end;
                2:begin                                             //班级
                  {登分表、word、单科、班级}
                  output1.stat_word_aCourse_class(report_classID_ComboBox.text);
                end;
              end;
            end;
          end;
        end;
        1:begin                                                    //excel
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {登分表、excel、全科、全校}
                  output1.stat_Excel_allCourse_school();
                end;
                1:begin                                            //全年级
                  {登分表、excel、全科、全年级}
                  output1.stat_Excel_allCourse_grade(report_grade_ComboBox.ItemIndex + 1);
                end;
                2:begin                                            //班级
                  {登分表、excel、全科、班级}
                  output1.stat_excel_allCourse_class(report_classID_ComboBox.text);
                end;
              end;
            end;
            1:begin                                                //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {登分表、excel、单科、全校}
                  output1.stat_Excel_aCourse_school();
                end;
                1:begin                                            //全年级
                  {登分表、excel、单科、全年级}
                  output1.stat_Excel_aCourse_grade(report_grade_ComboBox.ItemIndex + 1);
                end;
                2:begin                                            //班级
                  {登分表、excel、单科、班级}
                  output1.stat_excel_aCourse_class(report_classID_ComboBox.text);
                end;
              end;
            end;
          end;
        end;
      end;
    end;
    1: begin   //成绩表
      case report_Filetype_ComboBox.ItemIndex of
        0:begin                                                     //word
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                 //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {成绩表、word、全科、全校}
                end;
                1:begin                                             //全年级
                  {成绩表、word、全科、全年级}
                end;
                2:begin                                             //班级
                  {成绩表、word、全科、班级}
                end;
              end;
            end;
            1:begin                                                 //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {成绩表、word、单科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.score_word_aCourse_school(course , exam);
                end;
                1:begin                                             //全年级
                  {成绩表、word、单科、全年级}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  grade := report_grade_ComboBox.ItemIndex + 1;
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.score_word_aCourse_grade(grade, course , exam);
                end;
                2:begin                                             //班级
                  {成绩表、word、单科、班级}
                  classID := report_classID_ComboBox.text;
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.score_word_aCourse_class(classID, course , exam);
                end;
              end;
            end;
          end;
        end;
        1:begin                                                    //excel
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {成绩表、excel、全科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  exam := report_exam_ComboBox.Text;
                  output1.score_Excel_allCourse_school(exam);
                end;
                1:begin                                            //全年级
                  {成绩表、excel、全科、全年级}
                  grade := report_grade_ComboBox.ItemIndex + 1;
                  exam := report_exam_ComboBox.Text;
                  output1.score_Excel_allCourse_grade(grade, exam);
                end;
                2:begin                                            //班级
                  {成绩表、excel、全科、班级}
                  classID := report_classID_ComboBox.text;
                  exam := report_exam_ComboBox.Text;
                  output1.score_Excel_allCourse_class(classID,exam);
                end;
              end;
            end;
            1:begin                                                //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {成绩表、excel、单科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.score_Excel_aCourse_school(course, exam);
                end;
                1:begin                                            //全年级
                  {成绩表、excel、单科、全年级}
                  grade := report_grade_ComboBox.ItemIndex + 1;
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.score_Excel_aCourse_grade(grade, course , exam);
                end;
                2:begin                                            //班级
                  {成绩表、excel、单科、班级}
                  classID := report_classID_ComboBox.text;
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.score_Excel_aCourse_class(classID, course , exam);
                end;
              end;
            end;
          end;
        end;
      end;
    end;
    2: begin     //期末评估
      case report_Filetype_ComboBox.ItemIndex of
        0:begin                                                     //word
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                 //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {期末评估、word、全科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  exam := report_exam_ComboBox.Text;
                  output1.finalExam_word_allCourse_school(exam);
                end;
                1:begin                                             //全年级
                  {期末评估、word、全科、全年级}
                end;
                2:begin                                             //班级
                  {期末评估、word、全科、班级}
                end;
              end;
            end;
            1:begin                                                 //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {期末评估、word、单科、全校}
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.finalExam_word_aCourse_school(course, exam);
                end;
                1:begin                                             //全年级
                  {期末评估、word、单科、全年级}
                end;
                2:begin                                             //班级
                  {期末评估、word、单科、班级}
                end;
              end;
            end;
          end;
        end;
        1:begin                                                    //excel
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {期末评估、excel、全科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  exam := report_exam_ComboBox.Text;
                  output1.finalExam_Excel_allCourse_school(exam);
                end;
                1:begin                                            //全年级
                  {期末评估、excel、全科、全年级}
                end;
                2:begin                                            //班级
                  {期末评估、excel、全科、班级}
                end;
              end;
            end;
            1:begin                                                //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {期末评估、excel、单科、全校}
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.finalExam_Excel_aCourse_school(course, exam);
                end;
                1:begin                                            //全年级
                  {期末评估、excel、单科、全年级}
                end;
                2:begin                                            //班级
                  {期末评估、excel、单科、班级}
                end;
              end;
            end;
          end;
        end;
      end;
    end;
    3: begin   //考试评估
      case report_Filetype_ComboBox.ItemIndex of
        0:begin                                                     //word
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                 //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {考试评估、word、全科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  exam := report_exam_ComboBox.Text;
                  output1.exam_word_allCourse_school(exam);
                end;
                1:begin                                             //全年级
                  {考试评估、word、全科、全年级}
                end;
                2:begin                                             //班级
                  {考试评估、word、全科、班级}
                end;
              end;
            end;
            1:begin                                                 //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                             //全校
                  {考试评估、word、单科、全校}
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.exam_word_aCourse_school(course, exam);
                end;
                1:begin                                             //全年级
                  {考试评估、word、单科、全年级}
                end;
                2:begin                                             //班级
                  {考试评估、word、单科、班级}
                end;
              end;
            end;
          end;
        end;
        1:begin                                                    //excel
          case report_courseType_ComboBox.ItemIndex of
            0:begin                                                //全科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {考试评估、excel、全科、全校}
                  if visual_CheckBox.Checked then
                  begin
                    showmessage('多文件输出不支持预览。');
                    visual_CheckBox.Checked := false;
                    output1.isVisual(false);
                  end;
                  exam := report_exam_ComboBox.Text;
                  output1.exam_Excel_allCourse_school(exam);
                end;
                1:begin                                            //全年级
                  {考试评估、excel、全科、全年级}
                end;
                2:begin                                            //班级
                  {考试评估、excel、全科、班级}
                end;
              end;
            end;
            1:begin                                                //单科
              case report_outPutLimits_ComboBox.ItemIndex of
                0:begin                                            //全校
                  {考试评估、excel、单科、全校}
                  course := report_courseName_ComboBox.Text;
                  exam := report_exam_ComboBox.Text;
                  output1.exam_Excel_aCourse_school(course, exam);
                end;
                1:begin                                            //全年级
                  {考试评估、excel、单科、全年级}
                end;
                2:begin                                            //班级
                  {考试评估、excel、单科、班级}
                end;
              end;
            end;
          end;
        end;
      end;
    end;
  end;
//导出完成，消除提示信息
  promptPanel.SetBounds(Width,height,0,0);
  promptPanel.Visible := false;
end;


procedure TForm1.report_type_ComboBoxChange(Sender: TObject);
begin
  initOutputPanel();
  report_FileType_ComboBox.Enabled := true;
  report_courseType_ComboBox.Enabled := true;

  case report_type_ComboBox.ItemIndex of
    0: begin             //登分表
      report_outPutLimits_ComboBox.Enabled := true;
    end;
    1: begin             //成绩表
      report_outPutLimits_ComboBox.Enabled := true;
      report_exam_ComboBox.Enabled := true;          //report_exam_ComboBox中有所有考试
    end;
    2: begin             //期末评估
      report_courseName_ComboBox.Enabled := true;    //report_courseName_ComboBox中有所有年级课程
      report_outPutLimits_ComboBox.ItemIndex := 0;
      report_exam_ComboBox.Items.Clear;
      report_exam_ComboBox.Items.Add('期末');
      report_exam_ComboBox.ItemIndex := 0;
      report_exam_ComboBox.Enabled := false;
    end;
    3: begin             //考试评估
      report_exam_ComboBox.Enabled := true;          //report_exam_ComboBox中有所有考试
      report_courseName_ComboBox.Enabled := true;    //report_courseName_ComboBox中有所有年级课程
      report_outPutLimits_ComboBox.ItemIndex := 0;
    end;
  end;
end;




{***********************************文件导入***********************************}
procedure TForm1.input_getFile_ButtonClick(Sender: TObject);
begin
  if opendialog1.Execute then
    input_filePath_Edit.text:=opendialog1.FileName ;
end;

procedure TForm1.input_type_ComboBoxChange(Sender: TObject);
begin
  if input_type_ComboBox.ItemIndex = 3 then
  begin
    input_exam_ComboBox.Enabled := true;
  end else begin
    input_exam_ComboBox.Enabled := false;
  end;
end;


procedure TForm1.N2Click(Sender: TObject);
var
  form2 : TForm2;
begin
  Application.CreateForm(TForm2, Form2); { 创建 Form1 主窗体 }
  Form2.Show;                            { 显示 Form1，这时是主窗体了 }
  GlobalData.enterScoreIdentity := 0;   //表示从管理员进入成绩录入人员界面
end;



procedure TForm1.inout_submit_btnClick(Sender: TObject);
var
  path : string;
  input1 : Input;
begin
  promptPanel.SetBounds(round((Width-promptLabel.Width)/2),
      round((Height-promptLabel.Height)/2),promptLabel.Width,promptLabel.Height);
  promptLabel.SetBounds(0,0,promptPanel.Width,promptPanel.Height);
  promptLabel.Caption := '数据导入中，请稍后...';
  promptPanel.Visible := true;
  promptPanel.BringToFront;

  input1 := Input.Create;
  path := input_filePath_Edit.text;
  case input_type_ComboBox.ItemIndex of
    0:begin      //导入课程
      input1.course(path);
    end;
    1:begin      //导入班级
      input1.classes(path);
    end;
    2:begin      //导入学生
      input1.student(path);
    end;
    3:begin      //导入成绩
      if input_exam_ComboBox.ItemIndex = -1 then
      begin
        showmessage('请在考试类型选项中选择您本次导入的成绩属于哪次考试。');
      end else begin
        input1.scores(input_exam_ComboBox.Text,path);
      end;
    end;
  end;

  promptPanel.SetBounds(Width,height,0,0);
  promptPanel.Visible := false;
end;

{***********************************数据库管理*********************************}
procedure TForm1.term_ComboBoxChange(Sender: TObject);
begin
  if keepBack_CheckBox.Checked and (term_ComboBox.ItemIndex = 1) then            //下学期
  begin
    data_course_CheckBox.Enabled := true;
  end else begin
    data_course_CheckBox.Checked := false;
    data_course_CheckBox.Enabled := false;
  end;

end;

procedure TForm1.keepBack_CheckBoxClick(Sender: TObject);
begin
  if keepBack_CheckBox.Checked then
  begin
    if term_ComboBox.ItemIndex = 1 then
    begin
      data_course_CheckBox.Enabled := true;
    end else begin
      data_course_CheckBox.Checked := false;
      data_course_CheckBox.Enabled := false;
    end;
    data_par_CheckBox.Enabled := true;
    data_teacher_CheckBox.Enabled := true;
    data_user_CheckBox.Enabled := true;
    data_stu_CheckBox.Enabled := true;
  end else begin
    data_course_CheckBox.Enabled := false;
    data_par_CheckBox.Enabled := false;
    data_teacher_CheckBox.Enabled := false;
    data_user_CheckBox.Enabled := false;
    data_stu_CheckBox.Enabled := false;
  end;

end;

procedure TForm1.database_submit_btnClick(Sender: TObject);
var
  data : Database;
  flag : boolean;
  y : integer;
  I: Integer;

  term_ComboBox_index : integer;
  term : string;
begin
  data := Database.Create;
  flag := true;
  y:= check.year(database_beginYear_Edit.Text,flag);

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


  if flag then
  begin
    //显示提示信息
      promptPanel.SetBounds(round((Width-promptLabel.Width)/2),
      round((Height-promptLabel.Height)/2),promptLabel.Width,promptLabel.Height);
      promptLabel.SetBounds(0,0,promptPanel.Width,promptPanel.Height);
      promptLabel.Caption := '数据库初始化，请稍后...';
      promptPanel.Visible := true;
      promptPanel.BringToFront;
    //复制前一学期的数据库
    if not data.createDatabaseByOld(y,term) then
    begin
      showmessage('数据库创建失败！');
      exit;
    end;
    //保存原有数据库连接串
    data.saveConnString();
    //连接到新数据库
    if not data.linkNewDatabase(y,term) then
    begin
      showmessage('新数据库连接失败！');
      exit;
    end;
    //设置新数据库开学年份
    data.setBeginYear(y);

    //删除全部考试信息，以及对应成绩
    data.delExam(data_stu_CheckBox.Checked);

    //先处理课程和成绩信息
    if not data_course_CheckBox.Checked then
    begin
      data.delCourse();     //课程信息会被删除，对应成绩表也会删除
    end;

    //跨学年需要更新课程表
    if comparestr(term,'上学期') = 0 then
    begin
      data.updateCourseTable();
    end;
    //删除不需要保留的数据


    if not data_par_CheckBox.Checked then
      data.delPar();
    if not data_teacher_CheckBox.Checked then
      data.delTeacher();
    if not data_user_CheckBox.Checked then
      data.delUser();
    if not data_stu_CheckBox.Checked then
    begin
      for I := 1 to 9 do
      begin
        data.delStu(i);
      end;
    end;
    //恢复原有数据库连接
    data.recoverOldLink();
    //消除提示信息
    promptPanel.SetBounds(Width,height,0,0);
    promptPanel.Visible := false;
    showmessage('数据库创建完成');
  end;
  promptPanel.Visible := false;
end;
{*********************************全局辅助方法*********************************}
{
    根据选择信息界面动态变化
}
procedure TForm1.grade_ComboBoxChange(Sender: TObject);
var
  grade : integer;    //年级
  classID : string;
  ComboBox:TComboBox;
  trans : CTransForm;
  allGrade : boolean;
begin
  if PageControl1.ActivePage = coursePage then        //课程管理
  begin
    grade := course_grade_ComboBox.ItemIndex + 1;
    ComboBox := course_classID_ComboBox;
    getInfoByGrade(ComboBox,grade);
    course.getCourseByGrade(ADOQuery,grade);
  end;

  if PageControl1.ActivePage = studentPage then        //学生管理
  begin
    grade := stu_grade_ComboBox.ItemIndex + 1;
    ComboBox := stu_classID_ComboBox;
    getInfoByGrade(ComboBox,grade);
  end;

end;

procedure TForm1.getInfoByGrade(var ComboBox:TComboBox; grade:integer);
var
  ADOQuery1 : TADOQuery;
begin
  ADOQuery1 := TADOQuery.Create(nil);
// 查询年级为grade的所有班级
  classes.selectByGrade(ADOQuery1, grade);

//将查询得到的班级放入class_classID_ComboBox.Items中
  ComboBox.Items.Clear;
  while not ADOQuery1.eof do
  begin
    ComboBox.Items.Add(ADOQuery1.FieldByName('班号').AsString);
    ADOQuery1.Next;
  end;
end;

{
    根据表格中选中的数据动态填充控制区
}
procedure TForm1.DBGrid1CellClick(Column: TColumn);
var
  trans : CTransForm;
begin
  trans := CTransForm.Create;
  if PageControl1.ActivePage = classPage then        //班级管理
  begin
    if (class_operate_RG.ItemIndex = 1) or (class_operate_RG.ItemIndex = 2) then
    begin  //修改或删除
      class_classID_ComboBox.Text := ADOQuery.FieldByName('班号').AsString;
      class_studentsNum_Edit.Text := ADOQuery.FieldByName('学生人数').AsString;
      class_teacher_Edit.Text := ADOQuery.FieldByName('班主任').AsString;
      class_grade_ComboBox.ItemIndex := trans.classIDToGrade(ADOQuery.FieldByName('班号').AsString) - 1;
    end;
  end;

  if PageControl1.ActivePage = coursePage then        //课程管理
  begin
    if course_operate_RG.ItemIndex = 1 then      //修改
    begin
      try
        course_grade_ComboBox.ItemIndex := ADOQuery.FieldByName('年级').AsInteger - 1;
        course_courseName_Edit.Text := ADOQuery.FieldByName('课程名称').AsString;
      except
        course_courseName_Edit.Text := ADOQuery.FieldByName('课程名称').AsString;
        course_teacher_Edit.Text := ADOQuery.FieldByName('任课教师').AsString;
      end;
      if Comparestr('文科',ADOQuery.FieldByName('课程类型').AsString)=0 then
      begin
        course_courseType_ComboBox.ItemIndex := 1;
      end else if Comparestr('理科',ADOQuery.FieldByName('课程类型').AsString)=0 then
      begin
        course_courseType_ComboBox.ItemIndex := 0;
      end;

      getInfoByGrade(course_classID_ComboBox,course_grade_ComboBox.ItemIndex + 1);
    end;

    if course_operate_RG.ItemIndex = 2 then      //删除
    begin
      course_grade_ComboBox.ItemIndex := ADOQuery.FieldByName('年级').AsInteger - 1;
      course_courseName_Edit.Text := ADOQuery.FieldByName('课程名称').AsString;
      if Comparestr('文科',ADOQuery.FieldByName('课程类型').AsString)=0 then
      begin
        course_courseType_ComboBox.ItemIndex := 1;
      end else if Comparestr('理科',ADOQuery.FieldByName('课程类型').AsString)=0 then
      begin
        course_courseType_ComboBox.ItemIndex := 0;
      end;
    end;
  end;

  if PageControl1.ActivePage = studentPage then        //学生管理
  begin
    if (stu_operate_RG.ItemIndex = 1) or (stu_operate_RG.ItemIndex =2) then   //修改或者删除
    begin
      stu_stuID_Edit.Text := ADOQuery.FieldByName('学号').AsString;
      stu_name_Edit.Text := ADOQuery.FieldByName('姓名').AsString;
      stu_sex_ComboBox.Text := ADOQuery.FieldByName('性别').AsString;
      stu_age_ComboBox.Text := ADOQuery.FieldByName('年龄').AsString;
    end;
  end;

  if PageControl1.ActivePage = examPage then           //考务管理
  begin
    examName_Edit.Text := ADOQuery.FieldByName('考试名称').AsString;
    examTime_Edit.Text := ADOQuery.FieldByName('考试时间').AsString;
  end;

  if PageControl1.ActivePage = parPage then           //参数配置
  begin
    par_grade_ComboBox.ItemIndex  := ADOQuery.FieldByName('年级').AsInteger - 1;
    par_averWen_base_Edit.Text    := ADOQuery.FieldByName('文科平均分下界').AsString;
    par_averWen_top_Edit.Text     := ADOQuery.FieldByName('文科平均分上界').AsString;
    par_averLi_base_Edit.Text     := ADOQuery.FieldByName('理科平均分下界').AsString;
    par_averLi_top_Edit.Text      := ADOQuery.FieldByName('理科平均分上界').AsString;
    par_averWenNo_Edit.Text       := ADOQuery.FieldByName('文科均分否决').AsString;
    par_averWenYes_Edit.Text      := ADOQuery.FieldByName('文科均分肯定').AsString;
    par_averLiNo_Edit.Text        := ADOQuery.FieldByName('理科均分否决').AsString;
    par_averLiYes_Edit.Text       := ADOQuery.FieldByName('理科均分肯定').AsString;
    par_dev_Edit.Text             := ADOQuery.FieldByName('标准差评估差值').AsString;
    par_greatRatio_Edit.Text      := ADOQuery.FieldByName('优生率评估差值').AsString;
    par_greatWeight_Edit.Text     := ADOQuery.FieldByName('优生占比').AsString;
    par_inferiorWeight_Edit.Text  := ADOQuery.FieldByName('学困占比').AsString;
  end;

end;

end.
