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

    function isHasCourse(grade1 : integer; course : string):boolean;
    procedure getCourse(var ADOQuery : TADOQuery);
    function getAllCourse(var ADOQuery : TADOQuery):Integer;
    procedure getCourseByGrade(var ADOQuery : TADOQuery; grade : integer);
    procedure addCourse();
    procedure changeCourse();
    procedure delCourse();
end;

implementation
  Constructor Tb_course.create;
  begin
    ado := TAdoOperate.Create;
    classes := Tb_classes.Create;
  end;


  function Tb_course.isHasCourse(grade1 : integer; course : string):boolean;
  var
    sql : string;
    ADOQuery1 : TADOQuery;
  begin
    ADOQuery1 := TADOQuery.Create(nil);
    sql := 'SELECT courseName from  tb_course_'+inttostr(grade1)+' where courseName = '+#39+course+#39;
    ado.SelectInfo(ADOQuery1, sql);
    if ADOQuery1.RecordCount = 0 then
    begin
      result := false;
    end else begin
      result := true;
    end;
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
        'courseType as �γ�����, ' +
        'T_'+classID +' as �ον�ʦ ' +
        'from  tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
  end;

  function Tb_course.getAllCourse(var ADOQuery : TADOQuery):Integer;
  var
    I : Integer;
    count : integer;
    ADOQuery1 : TADOQuery;

    sql : string;

    grade : integer;
    courseName : string;
    courseType : string;

  begin
    ADOQuery1 := TADOQuery.Create(nil);
    try
      sql := 'Drop table Tb_courseTemp';
      ado.ExecSqlStr(sql);
    except

    end;

    sql := 'SELECT 1 as grade, courseName, courseType '+
        ' into Tb_courseTemp '+
        ' from tb_course_1';
    ado.ExecSqlStr(sql);

    for I := 2 to 9 do
    begin
      getCourseByGrade(ADOQuery1, i);
      while not ADOQuery1.Eof do
      begin
        grade := ADOQuery1.FieldByName('�꼶').AsInteger;
        courseName := ADOQuery1.FieldByName('�γ�����').AsString;
        courseType := ADOQuery1.FieldByName('�γ�����').AsString;
        sql := 'INSERT INTO Tb_courseTemp(grade, courseName, courseType) '+
            ' Values('   + inttostr(grade)      +','+
                     #39 + courseName      +#39 +','+
                     #39 + courseType      +#39 +')';
        ado.ExecSqlStr(sql);
        ADOQuery1.Next;
      end;
    end;

    sql := 'SELECT grade as �꼶, courseName as �γ�����,courseType as �γ����� '+
        ' from Tb_courseTemp';
    ado.SelectInfo(ADOQuery, sql);

    try
      sql := 'Drop table Tb_courseTemp';
      ado.ExecSqlStr(sql);
    except
    end;
  end;

  {
      ��ȡĳ���꼶�����пγ�
      ������  grade       �꼶
  }
  procedure Tb_course.getCourseByGrade(var ADOQuery : TADOQuery; grade : integer);
  var
    sql : string;
  begin
    sql := 'SELECT '+inttostr(grade)+' as �꼶, courseName as �γ�����,courseType as �γ�����,* '+
        ' from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      ��ĳ���꼶���һ�ſγ�
      ������  grade       �꼶
              CourseName   �γ���
              courseType    �γ�����
  }
  procedure Tb_course.addCourse();
  var
    sql : string;
    classIdList : TStringList;
    i : integer;
    exam : Tb_exam;
    examList : TStringList;
    ADOQuery : TADOQuery;
  begin
    ADOQuery := TADOQuery.Create(nil);
    exam := Tb_exam.Create;
    examList := TStringList.Create;
    classIdList := TStringList.Create;
    //�ж����ݿ����Ƿ���ڸÿγ�
    sql := 'select * from tb_course_'+inttostr(grade)+' where CourseName = '+#39+CourseName+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('���ݿ����Ѵ��ڿγ̣�'+inttostr(grade)+'�꼶:'+CourseName);
      exit;
    end;
    //����Ӧ�꼶�γ̱������һ����¼
    sql := 'INSERT INTO tb_course_'+inttostr(grade)+'(courseName, courseType) '+
        'Values(' + #39 + CourseName + #39 + ',' +
                    #39 + courseType + #39 + ')';
    ado.ExecSqlStr(sql);

    //Ϊ���꼶���¼���γ̴���һ�ųɼ���
    sql := 'Create TABLE tb_scores_' + inttostr(grade)+'_'+CourseName +'('+
        'stuID VarChar(10) primary key,'+
        'classID VarChar(5) not null,'+
        'stuName VarChar(10) not null,'+
        '��ĩ float default -1 )';
    ado.ExecSqlStr(sql);

    exam.selectExamName(grade,ADOQuery);
    while not ADOQuery.Eof do
    begin
      examList.Add(ADOQuery.FieldByName('��������').AsString );
      ADOQuery.Next;
    end;

    for I := 0 to examList.Count -1 do
    begin
      try
        sql := 'alter table tb_scores_' + inttostr(grade)+'_'+CourseName+
                ' add '+ examList [i] +' float default -1';
        ado.ExecSqlStr(sql);
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
      ado.ExecSqlStr(sql);
    end;
  end;


  {
      ����ĳ���༶ĳ�γ̵��ο���ʦ
      ������  grade       �꼶
              classID     ���
              tracher     �ο���ʦ����
              courseName  �γ�����
  }
  procedure Tb_course.changeCourse();
  var
    sql : string;
  begin
    sql := 'update tb_course_'+inttostr(grade)+' '+
        ' set T_'+classID+' ='+ #39 + teacher     + #39 +  ','
              +' courseType ='+ #39 + courseType  + #39 +
        ' where courseName=' + #39 + CourseName + #39;
    ado.ExecSqlStr(sql);
  end;

  {
      ɾ��ĳ���꼶��ĳ�γ�
      ������  grade       �꼶
              courseName  �γ�����
  }
  procedure Tb_course.delCourse();
  var
    sql : string;
  begin
    sql := 'DELETE FROM tb_course_' + inttostr(grade) +
        ' WHERE courseName ='+#39 + CourseName +#39;
    ado.ExecSqlStr(sql);
    //ɾ����γ���صı�
    sql := 'Drop table tb_scores_' + inttostr(grade)+'_'+CourseName;
    ado.ExecSqlStr(sql);
  end;

end.
