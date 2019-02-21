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
    examName : string;      //��������
    term : string;          //ѧ�ڣ�0��ʾ��ѧ�ڣ�1��ʾ��ѧ��
    examDate : string;      //����ʱ��
    examGrade: string;      //�ο��꼶

    procedure selectExam(var ADOQuery : TADOQuery);
    procedure selectExamName( grade:integer; term:string; var ADOQuery : TADOQuery);
    procedure addExam(var ADOQuery : TADOQuery);
    procedure changeExam(var ADOQuery : TADOQuery);
    procedure delExam(var ADOQuery : TADOQuery);
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
    sql := 'SELECT examName as ��������,'+
        ' term as ѧ��,'+
        ' examDate as ����ʱ��,'+
        ' examGrade as �ο��꼶'+
        ' from tb_exam'+
        ' order by examName asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      ͨ���꼶��ѧ����Ѱ�����ڸ��꼶��ѧ�ڵĿ���
  }
  procedure Tb_exam.selectExamName(grade: Integer; term:string; var ADOQuery: TADOQuery);
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

    sql := 'SELECT examName as �������� '+
        ' from tb_exam'+
        ' where term = '+#39+term+#39+' and examGrade like '+#39 + examOfGrade +#39+
        ' order by examName asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  procedure Tb_exam.addExam(var ADOQuery : TADOQuery);
  var
    sql : string;
    i : integer;
    j : integer;
    grade : integer;
    courseLsist : TStringList;
  begin
    sql := 'INSERT INTO tb_exam(examName, term, examDate, examGrade) '+
        ' Values(' + #39 + examName +#39 +','+
                     #39 + term     +#39 +','+
                     #39 + examDate +#39 +','+
                     #39 + examGrade+#39 +')';
    ado.ExecSqlStr(ADOQuery, sql);

    //����صķ���������Ӷ�Ӧ����
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
          //�ڶ�Ӧ�꼶�Ŀγ̵ķ����м���һ��
          sql := 'alter table tb_scores_'+inttostr(grade)+'_'+courseLsist[j]+
              ' add '+examName+'_'+term+' float default -1';
          ado.ExecSqlStr(ADOQuery, sql);
          sql := 'update tb_scores_'+inttostr(grade)+'_'+courseLsist[j]+
              ' set '+examName+'_'+term+' = -1';
          ado.ExecSqlStr(ADOQuery, sql);
        end;
      end;
    end;
  end;

  procedure Tb_exam.changeExam(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'update tb_exam set '+
        ' term= ' + #39 + term +#39 +','+
        ' examDate= ' + #39 + examDate +#39 +','+
        ' examGrade=' + #39 + examGrade+#39 +
        ' where examName=' + #39 + examName + #39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;

  procedure Tb_exam.delExam(var ADOQuery : TADOQuery);
  var
    sql : string;
    i : integer;
    j : integer;
    grade : integer;
    courseLsist : TStringList;
  begin
    //����صķ�������ɾ����Ӧ����
    sql := 'select examGrade from tb_exam '+
        ' where examName='+#39+examName+#39+
        ' and term='+#39+term+#39;
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
          //�ڶ�Ӧ�꼶�Ŀγ̵ķ�����ɾ��һ��
          sql := 'alter table tb_scores_'+inttostr(grade)+'_'+courseLsist[j]+
              ' DROP COLUMN '+examName+'_'+term;
          ado.ExecSqlStr(ADOQuery, sql);
        end;
      end;
    end;

    //ɾ�����Ա��ж�Ӧ�ļ�¼
    sql := 'DELETE FROM tb_exam' +
    ' WHERE examName ='+#39 + examName +#39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;
end.
