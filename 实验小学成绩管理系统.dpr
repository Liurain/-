program 实验小学成绩管理系统;

uses
  Vcl.Forms,
  AdminWin in 'AdminWin.pas' {Form1},
  ADOOperate in 'ADOOperate.pas',
  Table_classes in 'Table_classes.pas',
  Table_class in 'Table_class.pas',
  Table_course in 'Table_course.pas',
  Table_scores in 'Table_scores.pas',
  Table_assessment in 'Table_assessment.pas',
  CheckData in 'CheckData.pas',
  Table_parameters in 'Table_parameters.pas',
  EnterScoreWin in 'EnterScoreWin.pas' {Form2},
  LoginWin in 'LoginWin.pas' {Form3},
  Table_users in 'Table_users.pas',
  Table in 'Table.pas',
  DataTransform in 'DataTransform.pas',
  GlobalData in 'GlobalData.pas',
  Table_exam in 'Table_exam.pas',
  fileOutput in 'fileOutput.pas',
  fileInput in 'fileInput.pas',
  DatabaseManage in 'DatabaseManage.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm3, Form3);
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
