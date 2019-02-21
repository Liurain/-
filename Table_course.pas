unit Table_course;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Table_classes, Table_exam;

type
    Tb_course = class
    Constructor Create;
  private
     ado : TAdoOperate;
     classes : Tb_classes;
  public
    grade : integer;

    courseName : string;
    classID : string;
    courseType : string;  //�γ����(�Ŀơ����)
    teacher : string;  //�ον�ʦ

    procedure getCourse(var ADOQuery : TADOQuery);
    function getAllCourse(var courseList : TStringList;var gradeList : TStringList):Integer;
    procedure getCourseByGrade(var ADOQuery : TADOQuery; grade : integer);
    procedure addCourse(var ADOQuery : TADOQuery);
    procedure changeCourse(var ADOQuery : TADOQuery);
    procedure delCourse(var ADOQuery : TADOQuery);
end;

implementation
  Constructor Tb_course.create;
  begin
    ado := TAdoOperate.Create;
    classes := Tb_classes.Create;
  end;

  {
      ��ȡĳ���༶���пγ̵��ο����
      ������  classID     ���
              grade       �꼶
  }
  procedure Tb_course.getCourse(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT courseName as �γ�����, ' +
        'T_'+classID +' as �ον�ʦ ' +
        'from  tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
  end;

  function Tb_course.getAllCourse(var courseList : TStringList;var gradeList : TStringList):Integer;
  var
    I : Integer;
    count : integer;
    ADOQuery : TADOQuery;

  begin
    ADOQuery := TADOQuery.Create(nil);

    count := 0;
    for I := 1 to 9 do
    begin
      getCourseByGrade(ADOQuery, i);
      while not ADOQuery.Eof do
      begin
        courseList.Add(ADOQuery.FieldByName('�γ�����').AsString);
        gradeList.Add(inttostr(i));
        count := count + 1;
        ADOQuery.Next;
      end;
    end;
    result := count;
  end;

  {
      ��ȡĳ���꼶�����пγ�
      ������  grade       �꼶
  }
  procedure Tb_course.getCourseByGrade(var ADOQuery : TADOQuery; grade : integer);
  var
    sql : string;
  begin
    sql := 'SELECT courseName as �γ�����,courseType as �γ�����,* from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      ��ĳ���꼶���һ�ſγ�
      ������  grade       �꼶
              CourseName   �γ���
              courseType    �γ�����
  }
  procedure Tb_course.addCourse(var ADOQuery : TADOQuery);
  var
    sql : string;
    classIdList : TStringList;
    i : integer;
    exam : Tb_exam;
    examList : TStringList;
  begin
    exam := Tb_exam.Create;
    examList := TStringList.Create;
    classIdList := TStringList.Create;
    //�ж����ݿ����Ƿ���ڸÿγ�
    sql := 'select * from tb_course_'+inttostr(grade)+' where CourseName = '+#39+CourseName+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('���ݿ����Ѵ�γ̣�'+CourseName);
      exit;
    end;
    //����Ӧ�꼶�γ̱������һ����¼
    sql := 'INSERT INTO tb_course_'+inttostr(grade)+'(courseName, courseType) '+
        'Values(' + #39 + CourseName + #39 + ',' +
                    #39 + courseType + #39 + ')';
    ado.ExecSqlStr(ADOQuery, sql);

    //Ϊ���꼶���¼���γ̴���һ�ųɼ���
    sql := 'Create TABLE tb_scores_' + inttostr(grade)+'_'+CourseName +'('+
        'stuID VarChar(10) primary key,'+
        'classID VarChar(5) not null,'+
        'stuName VarChar(10) not null,'+
        '��ĩ_�� float default -1 ,'+
        '��ĩ_�� float default -1 )';
    ado.ExecSqlStr(ADOQuery, sql);
    exam.selectExamName(grade,'��',ADOQuery);
    while not ADOQuery.Eof do
    begin
      examList.Add(ADOQuery.FieldByName('��������').AsString + '_��');
      ADOQuery.Next;
    end;
    exam.selectExamName(grade,'��',ADOQuery);
    while not ADOQuery.Eof do
    begin
      examList.Add(ADOQuery.FieldByName('��������').AsString + '_��');
      ADOQuery.Next;
    end;

    for I := 0 to examList.Count -1 do
    begin
      try
        sql := 'alter table tb_scores_' + inttostr(grade)+'_'+CourseName+
                ' add '+ examList [i] +' float default -1';
        ado.ExecSqlStr(ADOQuery, sql);
      except

      end;
    end;



    //��ȡ��ǰ�꼶��ȫ���༶
    classes.selectByGrade(ADOQuery, grade);
    classIdList.Clear;
    while not ADOQuery.eof do
    begin
      classIdList.Add(ADOQuery.FieldByName('���').AsString);
      ADOQuery.Next;
    end;
    //�����꼶������ѧ�����뵽�ɼ�����
    for I := 0 to classIdList.Count - 1 do
    begin
      sql := 'Insert into tb_scores_' + inttostr(grade)+'_'+CourseName +
        '(stuID, stuName, classID) '+
        ' select stuID,stuName,'+#39+ classIdList[i]+#39 +'from tb_class_'+classIdList[i];
      ado.ExecSqlStr(ADOQuery, sql);
    end;
  end;


  {
      ����ĳ���༶ĳ�γ̵��ο���ʦ
      ������  grade       �꼶
              classID     ���
              tracher     �ο���ʦ����
              courseName  �γ�����
  }
  procedure Tb_course.changeCourse(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'update tb_course_'+inttostr(grade)+' '+
        ' set T_'+classID+' ='+ #39 + teacher     + #39 +  ','
              +' courseType ='+ #39 + courseType  + #39 +
        ' where courseName=' + #39 + CourseName + #39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;

  {
      ɾ��ĳ���꼶��ĳ�γ�
      ������  grade       �꼶
              courseName  �γ�����
  }
  procedure Tb_course.delCourse(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'DELETE FROM tb_course_' + inttostr(grade) +
        ' WHERE courseName ='+#39 + CourseName +#39;
    ado.ExecSqlStr(ADOQuery, sql);
    //ɾ����γ���صı�
    sql := 'Drop table tb_scores_' + inttostr(grade)+'_'+CourseName;
    ado.ExecSqlStr(ADOQuery, sql);
  end;

end.
