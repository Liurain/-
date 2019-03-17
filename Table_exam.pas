unit Table_exam;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, DataTransform, StrUtils;

type
    Tb_exam = class
    Constructor Create;
  private
    ado : TAdoOperate;
  public
    examName : string;      //考试名称
//    term : string;          //学期，0表示上学期，1表示下学期
    examDate : string;      //考试时间
    examGrade: string;      //参考年级

    procedure selectExam(var ADOQuery : TADOQuery);
    procedure selectExamName( grade:integer; var ADOQuery : TADOQuery);
    procedure addExam();
    procedure changeExam();
    procedure delExam(keepStu:boolean);
end;


implementation
  Constructor Tb_exam.create;
  begin
    ado := TAdoOperate.Create;
  end;

  procedure Tb_exam.selectExam(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT examName as 考试名称,'+
//        ' term as 学期,'+
        ' examDate as 考试时间,'+
        ' examGrade as 参考年级'+
        ' from tb_exam'+
        ' order by examName asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      通过年级和学期来寻找属于该年级该学期的考试
  }
  procedure Tb_exam.selectExamName(grade: Integer; var ADOQuery: TADOQuery);
  var
    examOfGrade : string;
    sql : string;
  begin
    case grade of
      1: examOfGrade := '________1';
      2: examOfGrade := '_______1_';
      3: examOfGrade := '______1__';
      4: examOfGrade := '_____1___';
      5: examOfGrade := '____1____';
      6: examOfGrade := '___1_____';
      7: examOfGrade := '__1______';
      8: examOfGrade := '_1_______';
      9: examOfGrade := '1________';
    end;

    sql := 'SELECT examName as 考试名称 '+
        ' from tb_exam'+
        ' where  examGrade like '+#39 + examOfGrade +#39+
        ' order by examName asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  procedure Tb_exam.addExam();
  var
    sql : string;
    i : integer;
    j : integer;
    grade : integer;
    courseLsist : TStringList;
    ADOQuery : TADOQuery;
  begin
    ADOQuery := TADOQuery.Create(nil);
    //判断数据库中是否存在该考试信息
    sql := 'select * from tb_exam where examName = '+
        #39+examName+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('数据库中已存在考试：'+examName);
      exit;
    end;

    //在相关的分数表中添加对应的列
    courseLsist := TStringList.Create;
    for I := 1 to 9 do
    begin
      if Comparestr(MidStr(examGrade,i,1), '1') = 0 then
      begin
        grade := 10 - i;
        sql := 'select courseName from tb_course_'+inttostr(grade);
        ado.SelectInfo(ADOQuery, sql);
        courseLsist.Clear;
        while not ADOQuery.eof do
        begin
          courseLsist.Add(ADOQuery.FieldByName('courseName').AsString);
          ADOQuery.Next;
        end;
        for j := 0 to courseLsist.Count-1 do
        begin
          //在对应年级的课程的分数中加入一列
          try
            sql := 'alter table tb_scores_'+inttostr(grade)+'_'+courseLsist[j]+
                ' add '+examName+' float default -1';
            ado.ExecSqlStr(sql);
            sql := 'update tb_scores_'+inttostr(grade)+'_'+courseLsist[j]+
                ' set '+examName+' = -1';
            ado.ExecSqlStr(sql);
          except
            showmessage(examName+'中包含不合法字符，如运算符 + - * /等等');
            exit;
          end;
        end;
      end;
    end;

    //向考试表中添加一行记录
    sql := 'INSERT INTO tb_exam(examName, examDate, examGrade) '+
        ' Values(' + #39 + examName +#39 +','+
                     #39 + examDate +#39 +','+
                     #39 + examGrade+#39 +')';
    ado.ExecSqlStr(sql);
  end;

  procedure Tb_exam.changeExam();
  var
    examDate1 : string;
    examGrade1 : string;
    examName1 : string;
  begin

    examDate1 := examDate;
    examGrade1 := examGrade;
    examName1 := examName;
    delExam(true);


    examDate := examDate1;
    examGrade := examGrade1;
    examName := examName1;
    addExam();
  end;



  procedure Tb_exam.delExam(keepStu:boolean);
  var
    sql : string;
    i : integer;
    j : integer;
    grade : integer;
    courseLsist : TStringList;
    ADOQuery : TADOQuery;
  begin
    ADOQuery := TADOQuery.Create(nil);
    //在相关的分数表中删除对应的列
    sql := 'select examGrade from tb_exam '+
        ' where examName='+#39+examName+#39;
    ado.SelectInfo(ADOQuery, sql);
    examGrade := ADOQuery.FieldByName('examGrade').AsString;
    courseLsist := TStringList.Create;
    for I := 1 to 9 do
    begin
      if Comparestr(MidStr(examGrade,i,1), '1') = 0 then
      begin
        grade := 10 - i;
        sql := 'select courseName from tb_course_'+inttostr(grade);
        ado.SelectInfo(ADOQuery, sql);
        courseLsist.Clear;
        while not ADOQuery.eof do
        begin
          courseLsist.Add(ADOQuery.FieldByName('courseName').AsString);
          ADOQuery.Next;
        end;
        for j := 0 to courseLsist.Count-1 do
        begin
          //在对应年级的课程的分数中删除一列
          sql := 'alter table tb_scores_'+inttostr(grade)+'_'+courseLsist[j]+
                ' DROP COLUMN '+examName;
          ado.ExecSqlStr(sql);
          if not keepStu then
          begin
            sql := 'delete from tb_scores_'+inttostr(grade)+'_'+courseLsist[j];
            ado.ExecSqlStr(sql);
          end;
        end;
      end;
    end;

    //删除考试表中对应的记录
    sql := 'DELETE FROM tb_exam' +
    ' WHERE examName ='+#39 + examName +#39;
    ado.ExecSqlStr(sql);
  end;
end.
