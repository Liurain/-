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
      获取某个班级所有课程的任课情况
      参数：  classID     班号
              grade       年级
  }
  procedure Tb_course.getCourse(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT courseName as 课程名称, ' +
        'courseType as 课程类型, ' +
        'T_'+classID +' as 任课教师 ' +
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
        grade := ADOQuery1.FieldByName('年级').AsInteger;
        courseName := ADOQuery1.FieldByName('课程名称').AsString;
        courseType := ADOQuery1.FieldByName('课程类型').AsString;
        sql := 'INSERT INTO Tb_courseTemp(grade, courseName, courseType) '+
            ' Values('   + inttostr(grade)      +','+
                     #39 + courseName      +#39 +','+
                     #39 + courseType      +#39 +')';
        ado.ExecSqlStr(sql);
        ADOQuery1.Next;
      end;
    end;

    sql := 'SELECT grade as 年级, courseName as 课程名称,courseType as 课程类型 '+
        ' from Tb_courseTemp';
    ado.SelectInfo(ADOQuery, sql);

    try
      sql := 'Drop table Tb_courseTemp';
      ado.ExecSqlStr(sql);
    except
    end;
  end;

  {
      获取某个年级的所有课程
      参数：  grade       年级
  }
  procedure Tb_course.getCourseByGrade(var ADOQuery : TADOQuery; grade : integer);
  var
    sql : string;
  begin
    sql := 'SELECT '+inttostr(grade)+' as 年级, courseName as 课程名称,courseType as 课程类型,* '+
        ' from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      向某个年级添加一门课程
      参数：  grade       年级
              CourseName   课程名
              courseType    课程类型
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
    //判断数据库中是否存在该课程
    sql := 'select * from tb_course_'+inttostr(grade)+' where CourseName = '+#39+CourseName+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('数据库中已存在课程：'+inttostr(grade)+'年级:'+CourseName);
      exit;
    end;
    //向相应年级课程表中添加一条记录
    sql := 'INSERT INTO tb_course_'+inttostr(grade)+'(courseName, courseType) '+
        'Values(' + #39 + CourseName + #39 + ',' +
                    #39 + courseType + #39 + ')';
    ado.ExecSqlStr(sql);

    //为该年级的新加入课程创建一张成绩表
    sql := 'Create TABLE tb_scores_' + inttostr(grade)+'_'+CourseName +'('+
        'stuID VarChar(10) primary key,'+
        'classID VarChar(5) not null,'+
        'stuName VarChar(10) not null,'+
        '期末 float default -1 )';
    ado.ExecSqlStr(sql);

    exam.selectExamName(grade,ADOQuery);
    while not ADOQuery.Eof do
    begin
      examList.Add(ADOQuery.FieldByName('考试名称').AsString );
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
      ado.ExecSqlStr(sql);
    end;
  end;


  {
      更换某个班级某课程的任课老师
      参数：  grade       年级
              classID     班号
              tracher     任课老师姓名
              courseName  课程名称
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
      删除某个年级的某课程
      参数：  grade       年级
              courseName  课程名称
  }
  procedure Tb_course.delCourse();
  var
    sql : string;
  begin
    sql := 'DELETE FROM tb_course_' + inttostr(grade) +
        ' WHERE courseName ='+#39 + CourseName +#39;
    ado.ExecSqlStr(sql);
    //删除与课程相关的表
    sql := 'Drop table tb_scores_' + inttostr(grade)+'_'+CourseName;
    ado.ExecSqlStr(sql);
  end;

end.
