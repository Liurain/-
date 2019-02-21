unit Table_class;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, DataTransform;

type
    Tb_class = class
    Constructor Create;
  private
    ado : TAdoOperate;      //数据库操作对象
    procedure changeScoreTable(flag : integer;var ADOQuery : TADOQuery);
  public
    classID : string;       //班号
    stuID :string;          //学号
    stuName : string;       //学生姓名
    sex : string;           //学生性别
    age : integer;          //学生年龄

    procedure getStu(var ADOQuery : TADOQuery);
    procedure addStu(var ADOQuery : TADOQuery);
    procedure changeStu(var ADOQuery : TADOQuery);
    procedure delStu(var ADOQuery : TADOQuery);
  end;

implementation
  Constructor Tb_class.create;
  begin
    ado := TAdoOperate.Create;
  end;

  {
      根据班级获取学生信息
      相关信息：
          学号    stuID
          姓名    stuName
          性别    stuSex
          年龄    stuAge
      预先设定参数：
          classID    班号
  }
  procedure Tb_class.getStu(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'SELECT stuID as 学号, '+
        ' stuName as 姓名, '+
        ' stuSex  as 性别, '+
        ' stuAge  as 年龄  '+
        ' from tb_class_'+classID+
        ' order by stuID asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      添加一个学生信息
      预先设定参数：
          classID    班号
          stuID      学号
          stuName    姓名
          sex        性别
          age        年龄
  }
  procedure Tb_class.addStu(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    //判断数据库中是否存在该学生
    sql := 'select * from tb_class_'+classID+' where stuID = '+#39+stuID+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('数据库中已存学生：'+stuID+'   '+stuName);
      exit;
    end;
    //向学生花名册中增加一条记录
    sql := 'INSERT INTO tb_class_'+classID+'(stuID,stuName,stuSex,stuAge) Values(' +
          #39 + stuID +#39 + ',' +
          #39 + stuName + #39 + ',' +
          #39 + sex + #39 +','+
          #39 + inttostr(age) + #39 +')';
    ado.ExecSqlStr(ADOQuery, sql);
    //该班级学生人数加一
    sql := 'update tb_classes set stuNum=stuNum + 1 where classID='+#39+classID+#39;
    ado.ExecSqlStr(ADOQuery, sql);
    //将学生所存在课程的课程表也增加一条记录
    changeScoreTable(0, ADOQuery);  //0表示增加一条记录
  end;


  {
      修改一个学生信息
      预先设定参数：
          classID    班号
          stuID      学号
          stuName    姓名
          sex        性别
          age        年龄
  }
  procedure Tb_class.changeStu(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'update tb_class_' + classID + ' set '+
        ' stuID   = ' + #39 + stuID   + #39 + ',' +
        ' stuName = ' + #39 + stuName + #39 + ',' +
        ' stuSex  = ' + #39 + sex     + #39 + ',' +
        ' stuAge  = ' + inttostr(age) +
        ' where stuID=' + #39 + stuID + #39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;


  {
      删除一个学生信息
      预先设定参数：
          stuID       学号
  }
  procedure Tb_class.delStu(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'DELETE FROM tb_class_' + classID +
        ' WHERE stuID = '+#39+stuID+#39;
    ado.ExecSqlStr(ADOQuery, sql);
    //该班级学生人数减一
    sql := 'update tb_classes set stuNum=stuNum - 1 where classID='+#39+classID+#39;
    ado.ExecSqlStr(ADOQuery, sql);

    //将学生所存在课程的课程表也删除一条记录
    changeScoreTable(1, ADOQuery);  //1表示删除一条记录
  end;

  procedure Tb_class.changeScoreTable(flag : integer;var ADOQuery : TADOQuery);
  var
    sql : string;
    courseList : TStringList;
    transform : CTransform;
    grade : integer;
    i : integer;
  begin
    courseList := TStringList.Create;
    transform := CTransform.Create;
    grade := transform.classIDToGrade(classID);
    sql := 'select courseName from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
    while not ADOQuery.eof do
    begin
      courseList.Add(ADOQuery.FieldByName('courseName').AsString);
      ADOQuery.Next;
    end;

    for I := 0 to courseList.Count - 1 do
    begin
      case flag of
        0:begin      //增加一个记录
          sql := 'INSERT INTO tb_scores_'+inttostr(grade)+'_'+courseList[i]+'(stuID,classID,stuName) Values(' +
                    #39 + stuID   + #39 + ',' +
                    #39 + classID + #39 + ',' +
                    #39 + stuName + #39 + ')';
        end;
        1:begin      //删除一条记录
          sql := 'DELETE FROM tb_scores_'+inttostr(grade)+'_'+courseList[i]+
                    ' WHERE stuID = '+#39+stuID+#39;
        end;
      end;

      ado.ExecSqlStr(ADOQuery, sql);
    end;
  end;

end.
