unit Table_classes;

interface

uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, DataTransform;

type
    Tb_classes = class
    Constructor Create;
  private
    ado : TAdoOperate;
  public
    grade : integer;      //年级
    classID: string;      //班号
    stuNum: integer;      //学生人数
    teacher: string;      //班主任

    procedure selectByGrade(var ADOQuery : TADOQuery; grade1 : integer);
    procedure selectAll(var ADOQuery : TADOQuery);

    procedure addClass();
    procedure changeClass();
    procedure delClass();
    function getStuNum(classID : string):integer;
end;


implementation
  Constructor Tb_classes.create;
  begin
    ado := TAdoOperate.Create;
  end;

  procedure Tb_classes.selectAll(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT classID as 班号, '+
        ' stuNum as 学生人数, '+
        ' teacher as 班主任 '+
        ' from tb_classes '+
        ' order by classID asc';
    ado.SelectInfo(ADOQuery, sql);
  end;


  {
      根据年级获取该年级相关信息
      参数：grade --年级
  }
  procedure Tb_classes.selectByGrade(var ADOQuery : TADOQuery; grade1 : integer);
  var
    sql : string;
    year : string; //学号中的年份
    transform : CTransform;
  begin
    transform := CTransform.Create;
    year := inttostr(transform.GradeToClassID(grade1));
    sql := 'SELECT classID as 班号, '+
        ' stuNum as 学生人数, '+
        ' teacher as 班主任 '+
        ' from tb_classes '+
        ' where classID like '+ year +'& "%"'+
        ' order by classID asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  function Tb_classes.getStuNum(classID : string):integer;
  var
    sql : string;
    ADOQuery1 : TADOQuery;
  begin
    ADOQuery1 := TADOQuery.Create(nil);
    sql := 'SELECT stuNum  from tb_classes '+
        ' where classID = '+#39+classID+#39;
    ado.SelectInfo(ADOQuery1, sql);
    result := ADOQuery1.FieldByName('stuNum').AsInteger;
  end;

  {
      向数据库中添加一个班级
      参数：
            classID   --班号，例如01
            stuNum    --本班级学生人数
            teacher   --班主任
            grade     --年级
  }
  procedure Tb_classes.addClass();
  var
    sql :string;
    ADOQuery : TADOQuery;
  begin
    ADOQuery := TADOQuery.Create(nil);
    //判断数据库中是否存在改班级
    sql := 'select * from tb_classes where classID = '+#39+classID+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('数据库中已存在'+classID+'班');
      exit;
    end;

    //向tb_classes表中添加一条班级信息
    if strtoint(classID) < 1000 then
    begin
      classID := '0'+ inttostr(strtoint(classID));
    end;
    sql := 'INSERT INTO tb_classes(classID,stuNum,teacher) Values(' +
            #39 + classID + #39 + ',' +
            #39 + inttostr(stuNum) + #39 +',' +
            #39 + teacher + #39 +')';
    ado.ExecSqlStr(sql);

    //为该班级创建一张花名册表
    sql := 'Create TABLE tb_class_' + classID +'('+
        'stuID VarChar(10) primary key,'+
        'stuName VarChar(10) not null,'+
        'stuSex VarChar(5) default '+ #39 +'男'+ #39+','+
        'stuAge int default 0 )';
    ado.ExecSqlStr(sql);

    //在对应年级的课程表中加入一列
    sql := 'alter table tb_course_'+inttostr(grade)+
        ' add T_'+classID+' VarChar(10)  null';
    ado.ExecSqlStr(sql);
  end;


  {
      修改数据库中一个班级的信息
      参数：grade         --班级所在年级
            classID_str   --班号，例如01
            stuNum_str    --本班级学生人数
            teacher_str   --班主任
  }
  procedure Tb_classes.changeClass();
  var
    sql :string;
  begin
    sql := 'update tb_classes set '+
        ' stuNum = ' + inttostr(stuNum) + ',' +
        ' teacher =' + #39 + teacher + #39 +
        ' where classID=' + #39 + classID + #39;
    ado.ExecSqlStr(sql);
  end;


  {
      删除数据库中一个班级的信息
      参数：grade         --班级所在年级
            classID_str   --班号，例如01
            stuNum_str    --本班级学生人数
            teacher_str   --班主任
  }
  procedure Tb_classes.delClass();
  var
    sql :string;
  begin
    //从班级表中删除班级信息
    sql := 'DELETE FROM tb_classes WHERE classID = '+#39+classID+#39;
    ado.ExecSqlStr(sql);
    //删除该班级相关的表
    sql := 'Drop table tb_class_'+classID;
    ado.ExecSqlStr(sql);
    //删除课程表中相关列
    sql := 'ALTER TABLE tb_course_'+inttostr(grade)+' DROP COLUMN T_'+classID;
    ado.ExecSqlStr(sql);
  end;



end.
