unit Table_scores;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate;
type
    Tb_scores = class
    Constructor create;
  private
    ado : TAdoOperate;      //���ݿ��������
  public
    courseName : string;
    grade : integer;
    stuID : string;
    stuName : string;
    score : real;
    classID : string;
    exam : string;          //��������

    procedure getScores(var ADOQuery : TADOQuery);
    procedure getScoresOfExam(orderType : string; var ADOQuery : TADOQuery);
    procedure setScores(var ADOQuery : TADOQuery);
end;

implementation
  Constructor Tb_scores.create;
  begin
    ado := TAdoOperate.Create;
  end;

  procedure Tb_scores.getScores(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'SELECT stuID as ѧ��, '+
        ' stuName  as ����, '+
        ' classID as ���, *'+
        ' from tb_scores_'+inttostr(grade)+'_'+courseName+
        ' where classID = '+ #39 + classID + #39+
        ' order by stuID asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  procedure Tb_scores.getScoresOfExam(orderType : string; var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'SELECT stuID as ѧ��, '+
        ' stuName  as ����, '+
        ' '+exam+' as ����'+
        ' from tb_scores_'+inttostr(grade)+'_'+courseName+
        ' where classID = '+ #39 + classID + #39+
        ' order by '+ orderType +' asc';
    ado.SelectInfo(ADOQuery, sql);

  end;

  procedure Tb_scores.setScores(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'update tb_scores_'+inttostr(grade)+'_'+courseName+ ' set '+
        exam+' = ' + floattostr(score)  +
        ' where stuID=' + #39 + stuID + #39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;
end.
