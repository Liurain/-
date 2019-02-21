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
    courseType : string;  //课程类别(文科、理科)
    teacher : string;  //任课教师

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
      获取某个班级所有课程的任课情况
      参数：  classID     班号
              grade       年级
  }
  procedure Tb_course.getCourse(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT courseName as 课程名称, ' +
        'T_'+classID +' as 任课教师 ' +
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
        courseList.Add(ADOQuery.FieldByName('课程名称').AsString);
        gradeList.Add(inttostr(i));
        count := count + 1;
        ADOQuery.Next;
      end;
    end;
    result := count;
  end;

  {
      获取某个年级的所有课程
      参数：  grade       年级
  }
  procedure Tb_course.getCourseByGrade(var ADOQuery : TADOQuery; grade : integer);
  var
    sql : string;
  begin
    sql := 'SELECT courseName as 课程名称,courseType as 课程类型,* from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      向某个年级添加一门课程
      参数：  grade       年级
              CourseName   课程名
              courseType    课程类型
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
    //判断数据库中是否存在该课程
    sql := 'select * from tb_course_'+inttostr(grade)+' where CourseName = '+#39+CourseName+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('数据库中已存课程：'+CourseName);
      exit;
    end;
    //向相应年级课程表中添加一条记录
    sql := 'INSERT INTO tb_course_'+inttostr(grade)+'(courseName, courseType) '+
        'Values(' + #39 + CourseName + #39 + ',' +
                    #39 + courseType + #39 + ')';
    ado.ExecSqlStr(ADOQuery, sql);

    //为该年级的新加入课程创建一张成绩表
    sql := 'Create TABLE tb_scores_' + inttostr(grade)+'_'+CourseName +'('+
        'stuID VarChar(10) primary key,'+
        'classID VarChar(5) not null,'+
        'stuName VarChar(10) not null,'+
        '期末_上 float default -1 ,'+
        '期末_下 float default -1 )';
    ado.ExecSqlStr(ADOQuery, sql);
    exam.selectExamName(grade,'上',ADOQuery);
    while not ADOQuery.Eof do
    begin
      examList.Add(ADOQuery.FieldByName('考试名称').AsString + '_上');
      ADOQuery.Next;
    end;
    exam.selectExamName(grade,'下',ADOQuery);
    while not ADOQuery.Eof do
    begin
      examList.Add(ADOQuery.FieldByName('考试名称').AsString + '_下');
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



    //获取当前年级的全部班级
    classes.selectByGrade(ADOQuery, grade);
    classIdList.Clear;
    while not ADOQuery.eof do
    begin
      classIdList.Add(ADOQuery.FieldByName('班号').AsString);
      ADOQuery.Next;
    end;
    //将该年级的所有学生导入到成绩表中
    for I := 0 to classIdList.Count - 1 do
    begin
      sql := 'Insert into tb_scores_' + inttostr(grade)+'_'+CourseName +
        '(stuID, stuName, classID) '+
        ' select stuID,stuName,'+#39+ classIdList[i]+#39 +'from tb_class_'+classIdList[i];
      ado.ExecSqlStr(ADOQuery, sql);
    end;
  end;


  {
      更换某个班级某课程的任课老师
      参数：  grade       年级
              classID     班号
              tracher     任课老师姓名
              courseName  课程名称
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
      删除某个年级的某课程
      参数：  grade       年级
              courseName  课程名称
  }
  procedure Tb_course.delCourse(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'DELETE FROM tb_course_' + inttostr(grade) +
        ' WHERE courseName ='+#39 + CourseName +#39;
    ado.ExecSqlStr(ADOQuery, sql);
    //删除与课程相关的表
    sql := 'Drop table tb_scores_' + inttostr(grade)+'_'+CourseName;
    ado.ExecSqlStr(ADOQuery, sql);
  end;

end.
