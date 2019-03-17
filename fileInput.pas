unit fileInput;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Vcl.OleServer, Word2000, Comobj, StrUtils,
  Table_course, Table_class, Table_scores, Table_classes, DataTransform;

type
    Input = class
    Constructor Create;
  private
    ado : TAdoOperate;        //数据库操作对象
    ADOQuery : TADOQuery;
    softPath : string;

    procedure createExcel(path : string; var ExcelApp: OleVariant);
  public
    procedure course(path : string);
    procedure student(path : string);
    procedure scores(exam : string; path : string);
    procedure classes(path : string);
  end;
var
  ExcelApp: OleVariant;

implementation
  Constructor Input.Create;
  begin
    ado := TAdoOperate.Create;
    ADOQuery := TADOQuery.Create(nil);
    softPath := ExtractFileDir(ParamStr(0));      //获取软件本身所在路径，用于文件操作

    try
      ExcelApp.Quit;       //关闭上一次预览未关闭的窗体并释放文件访问
    except
    end;
  end;

  //导入课程
  procedure Input.course(path : string);
  var
    sql : string;
//    ExcelApp: OleVariant;
    course : Tb_course;
    rowsCount : integer;
    I: Integer;
  begin
    course := Tb_course.Create;
    createExcel(path, ExcelApp);
    rowsCount := ExcelApp.ActiveSheet.UsedRange.Rows.Count;

    for I := 2 to rowsCount do        //第一行为表头
    begin
      course.grade := ExcelApp.ActiveSheet.Cells[i,1];
      course.courseName := ExcelApp.ActiveSheet.Cells[i,2];
      course.courseType := ExcelApp.ActiveSheet.Cells[i,3];
      try
        course.addCourse();
      except

      end;
    end;

    ExcelApp.quit;
  end;

  //导入学生
  procedure Input.student(path : string);
  var
    sql : string;
//    ExcelApp: OleVariant;
    student : Tb_class;
    classes : Tb_classes;
    rowsCount : integer;
    classIDList : TStringList;
    newClassIDList : TStringList;   //在导入学生的过程中新增加的班级
    trans : CTransform;

    classID : string;
    i : integer;
  begin
    student := Tb_class.Create;
    classes := Tb_classes.Create;
    classIDList := TStringList.Create;
    newClassIDList := TStringList.Create;
    trans := CTransform.Create;
    //获取数据库中所有班级号，用于判断班级是否存在
    classes.selectAll(ADOQuery);
    while not ADOQuery.Eof do
    begin
      classIDList.Add(ADOQuery.FieldByName('班号').AsString);
      ADOQuery.Next;
    end;

    createExcel(path, ExcelApp);
    rowsCount := ExcelApp.ActiveSheet.UsedRange.Rows.Count;
//     ExcelApp.Selection.Sort(
//          Key1 := ExcelApp.ActiveSheet.Range['A'+IntTostr(HbRowNum)],
//          Order1 := xlAscending);

    for I := 2 to rowsCount do        //第一行为表头
    begin
      student.stuID := ExcelApp.ActiveSheet.Cells[i,1];
      student.stuName := ExcelApp.ActiveSheet.Cells[i,2];
      student.sex := ExcelApp.ActiveSheet.Cells[i,3];
      student.age := ExcelApp.ActiveSheet.Cells[i,4];
      student.classID := LeftStr(student.stuID,4);

      if comparestr(classID,student.classID) <> 0 then
      begin
        classID := student.classID;
        if classIDList.IndexOf(classID) = -1 then       //不存在这个班级
        begin
          classes.classID := student.classID;
          classes.stuNum := 0;
          classes.teacher := '';
          classes.grade := trans.classIDToGrade(classes.classID);
          if newClassIDList.IndexOf(classID) = -1 then
          begin
            try
              classes.addClass();
              newClassIDList.Add(classes.classID);
            except
            end;
          end;
        end else begin
          sql := 'delete * from tb_class_'+classID;
          ado.ExecSqlStr(sql);
          sql := 'update tb_classes set stuNum = 0 where classID=' + #39 + classID + #39;;
          ado.ExecSqlStr(sql);
        end;
      end;

      try
        student.addStu();
      except

      end;
    end;

    ExcelApp.quit;
  end;

  procedure Input.scores(exam : string; path : string);
  var
    sql : string;
//    ExcelApp: OleVariant;
    scores : Tb_scores;
    rowsCount : integer;
    columns : integer;
    i : integer;     //表格的行计数变量
    j : integer;     //表格的列计数变量

    stuID : string;
    classID : string;
    grade : integer;
    trans : CTransform;

    course : Tb_Course;
    courseName : string;
    allCourseList : array[1..9] of TStringList;    //存储所有年级课程，用于判断课程是否开设
  begin
    scores := Tb_scores.create;
    trans := CTransform.Create;
    course := Tb_Course.Create;

    createExcel(path, ExcelApp);
    rowsCount := ExcelApp.ActiveSheet.UsedRange.Rows.Count;
    columns:= ExcelApp.ActiveSheet.UsedRange.columns.count;

    //初始化allCourseList
    for I := 1 to 9 do
    begin
      allCourseList[i] := TStringList.Create;
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        allCourseList[i].Add(ADOQuery.FieldByName('课程名称').AsString);
        ADOQuery.Next;
      end;
    end;

    for j := 3 to columns do          //第一列为学号，第二列为姓名
    begin
      for I := 2 to rowsCount do        //第一行为表头
      begin
        //从学号分析出该学生所在年级
        stuID := ExcelApp.ActiveSheet.Cells[i,1];
        classID := LeftStr(stuID,4);
        grade := trans.classIDToGrade(classID);
        //判断该年级是否有开设这个课程
        courseName := ExcelApp.ActiveSheet.Cells[1,j];
        if allCourseList[grade].IndexOf(courseName) <> -1 then      //有这门课
        begin
          scores.grade := grade;
          scores.courseName := courseName;
          scores.exam := exam;
          scores.stuID := ExcelApp.ActiveSheet.Cells[i,1];
          scores.score := ExcelApp.ActiveSheet.Cells[i,j];
          try
            scores.setScores();
          except

          end;
        end else begin
          //输出提示信息
        end;
      end;
    end;

    ExcelApp.quit;
  end;

  procedure Input.classes(path : string);
  var
    sql : string;
//    ExcelApp: OleVariant;
    classes : Tb_classes;
    rowsCount : integer;
    trans : CTransform;

    i : integer;
  begin
    classes := Tb_classes.Create;
    trans := CTransform.Create;
    createExcel(path, ExcelApp);
    rowsCount := ExcelApp.ActiveSheet.UsedRange.Rows.Count;

    for I := 2 to rowsCount do        //第一行为表头
    begin
      classes.classID := ExcelApp.ActiveSheet.Cells[i,1];
      classes.stuNum := 0;
      classes.teacher := ExcelApp.ActiveSheet.Cells[i,2];
      classes.grade := trans.classIDToGrade(classes.classID);
      try
        classes.addClass();
      except

      end;
    end;

    ExcelApp.quit;
  end;

  procedure Input.createExcel(path : string; var ExcelApp: OleVariant);
  begin
//    try
//      ExcelApp := GetActiveOleObject('Excel.Application');
//    except
      try
        ExcelApp := CreateOleObject('Excel.Application');
        ExcelApp.Visible := false;
      except
        ShowMessage('创建excel对象失败！请确认您使用的电脑是否安装有excel');
        Exit;
      end;
//    end;

      try
        ExcelApp.WorkBooks.Open(path);
      except
        ShowMessage('Excel文件'+path+'无法读取');
        Exit;
      end;

  end;


end.
