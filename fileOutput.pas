unit fileOutput;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Vcl.OleServer, Word2000, Comobj,DataTransform,
  StrUtils,GlobalData,Table_classes, Table_course,Table_assessment ;

type
    Output = class
    Constructor Create;
  private
    ado : TAdoOperate;        //数据库操作对象
    ADOQuery : TADOQuery;
    softPath : string;
    visual : boolean; //是否预览
    filePrint : boolean; //是否打印
    wordAppCanClose : boolean;       //wordApp能否关闭，用于判断是否追加数据
    excelAppCanClose : boolean;      //excelApp能否关闭，用于判断是否追加数据



    procedure createWord(modelFile, path, fileName : string;
        var wordApp, wordDoc, table : OleVariant; Columns : integer);
    procedure createExcel(modelFile, path, fileName : string;var ExcelApp : OleVariant);
    function addScoresExcel(classID, exam : string; var courseList : TstringList;
        rowIndex : integer; var ExcelApp : OleVariant):integer;
    function insterClassStuInfoToExcel(index : integer; classID : string):integer;
    function insterGradeStuInfoToExcel(index : integer; grade : integer):integer;
    function insterSchoolStuInfoToExcel(index : integer):integer;

    procedure FianlExamDataToWord(exam : string; course : string; var table: OleVariant);
    procedure secondTermFianlExamDataToWord(exam : string; course : string; var table: OleVariant);
    procedure FianlExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
    procedure secondTermFianlExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
    procedure ExamDataToWord(exam : string; course : string; var table: OleVariant);
    procedure ExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
  public
    procedure finalReport(course : string);
    procedure examReport(examName, course : string);
    procedure roster(classID : string;course,exam : string);
    procedure rosterExcel(classID : string;course,exam : string);
    procedure allScoresExcel(classID : string;exam : string);

    procedure isVisual(flag : boolean);
    procedure isfilePrint(flag : boolean);

{-------------------------------成绩统计表方法----------------------------------}
    procedure stat_word_allCourse_school();
    procedure stat_word_allCourse_grade(grade : integer);
    procedure stat_word_allCourse_class(classID : string);
    procedure stat_word_aCourse_school();
    procedure stat_word_aCourse_grade(grade : integer);
    procedure stat_word_aCourse_class(classID : string);
    procedure stat_Excel_allCourse_school();
    procedure stat_Excel_allCourse_grade(grade : integer);
    procedure stat_Excel_allCourse_class(classID : string);
    procedure stat_Excel_aCourse_school();
    procedure stat_Excel_aCourse_grade(grade : integer);
    procedure stat_Excel_aCourse_class(classID : string);
{---------------------------------成绩表方法------------------------------------}
    procedure score_word_allCourse_school();
    procedure score_word_allCourse_grade();
    procedure score_word_allCourse_class();
    procedure score_word_aCourse_school(course : string; exam : string);
    procedure score_word_aCourse_grade(grade : integer; course : string; exam : string);
    procedure score_word_aCourse_class(classID : string; course : string; exam : string);
    procedure score_Excel_allCourse_school(exam : string);
    procedure score_Excel_allCourse_grade(grade : integer; exam : string);
    procedure score_Excel_allCourse_class(classID : string; exam : string);
    procedure score_Excel_aCourse_school(course : string; exam : string);
    procedure score_Excel_aCourse_grade(grade : integer; course : string; exam : string);
    procedure score_Excel_aCourse_class(classID : string; course : string; exam : string);
{-------------------------------期末评估表方法----------------------------------}
    procedure finalExam_word_allCourse_school(exam : string);
//    procedure finalExam_word_allCourse_grade();
//    procedure finalExam_word_allCourse_class();
    procedure finalExam_word_aCourse_school(course : string; exam : string);
//    procedure finalExam_word_aCourse_grade();
//    procedure finalExam_word_aCourse_class();
    procedure finalExam_Excel_allCourse_school(exam : string);
//    procedure finalExam_Excel_allCourse_grade();
//    procedure finalExam_Excel_allCourse_class();
    procedure finalExam_Excel_aCourse_school(course : string; exam : string);
//    procedure finalExam_Excel_aCourse_grade();
//    procedure finalExam_Excel_aCourse_class();
{-------------------------------考试评估表方法----------------------------------}
    procedure exam_word_allCourse_school(exam : string);
//    procedure exam_word_allCourse_grade();
//    procedure exam_word_allCourse_class();
    procedure exam_word_aCourse_school(course : string; exam : string);
//    procedure exam_word_aCourse_grade();
//    procedure exam_word_aCourse_class();
    procedure exam_Excel_allCourse_school(exam : string);
//    procedure exam_Excel_allCourse_grade();
//    procedure exam_Excel_allCourse_class();
    procedure exam_Excel_aCourse_school(course : string; exam : string);
//    procedure exam_Excel_aCourse_grade();
//    procedure exam_Excel_aCourse_class();
  end;
var
  wordApp: OleVariant;
  ExcelApp: OleVariant;

implementation
  Constructor Output.Create;
  begin
    ado := TAdoOperate.Create;
    ADOQuery := TADOQuery.Create(nil);
    softPath := ExtractFileDir(ParamStr(0));      //获取软件本身所在路径，用于文件操作
    visual := false;

    try
//      while True do
//      begin
//        wordApp := GetActiveOleObject('Word.Application');
        wordApp.Quit;           //关闭上一次预览未关闭的窗体并释放文件访问
//      end;
    except
    end;

    try
//      while True do
        ExcelApp.Quit;
    except
    end;

  end;

  procedure Output.isVisual(flag : boolean);
  begin
    visual := flag;
  end;

  procedure Output.isfilePrint(flag : boolean);
  begin
    filePrint := flag;
  end;


{-------------------------------成绩统计表方法----------------------------------}
  procedure Output.stat_word_allCourse_school();
  var
    sql : string;
    fileName : string;
  begin

  end;

  procedure Output.stat_word_allCourse_grade(grade : integer);
  var
    sql : string;
  begin

  end;

  procedure Output.stat_word_allCourse_class(classID : string);
  begin

  end;
//  var
//    sql : string;
//    course : Tb_course;
//    trans : CTransform;
//    grade : integer;
//    i : integer;
//
//    wordDoc: OleVariant;
//    table: OleVariant;
//    modelFile : string;     //模板文件
//    path : string;          //文件输出路径
//    fileName : string;      //报表文件名
//    j : integer;
//    dataRows : integer;
//
//    Present: TDateTime;
//    Year, Month, Day :Word;
//    titleInfo :string;
//  begin
//  //复制模版并创建word对象打开模板
//    modelFile := softPath + '\model\全科登分表.docx';
//    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
//        '学年\登分表\全科登分表\Word文件';
//    fileName:= path + '\' + classID + '全科登分表.docx';
//    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //学生成绩统计表模板9列
//  //获取这个年级所有课程
//    course := Tb_course.Create;
//    trans := CTransform.Create;
//    grade := trans.classIDToGrade(classID);
//    course.getCourseByGrade(ADOQuery, grade);
//    i := 3;   //从第三列开始
//    while not ADOQuery.eof do
//    begin
//      ExcelApp.ActiveSheet.Cells[1,i] := ADOQuery.FieldByName('课程名称').AsString;
//      i := i+1;
//      ADOQuery.Next;
//    end;
//  end;

  procedure Output.stat_word_aCourse_school();
  var
    I : Integer;
  begin
    for I := 1 to 9 do
    begin
      stat_word_aCourse_grade(i);
    end;
  end;

  procedure Output.stat_word_aCourse_grade(grade : integer);
  var
    sql : string;
    ADOQuery1 : TADOQuery;
    classes : Tb_classes;
  begin
    ADOQuery1 := TADOQuery.Create(nil);
    classes := Tb_classes.Create;
  //获取这个年级所有的班级
    classes.selectByGrade(ADOQuery1, grade);
    while not ADOQuery1.Eof do
    begin
      stat_word_aCourse_class(ADOQuery1.FieldByName('班号').AsString);
      ADOQuery1.Next;
    end;
  end;

  procedure Output.stat_word_aCourse_class(classID : string);
  var
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    path : string;          //文件输出路径
    fileName : string;      //报表文件名
    i : integer;
    j : integer;
    dataRows : integer;
    sql : string;

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
  //复制模版并创建word对象打开模板
    modelFile := softPath + '\model\单科登分表.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\登分表\单科登分表\Word文件';
    fileName:= path + '\' + classID + '单科登分表.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //学生成绩统计表模板9列

  //加入表头信息
    trans := CTransform.Create;
    Present:= Now;
    DecodeDate(Present, Year, Month, Day);
    titleInfo := '班级:'+trans.classIDToText(classID)+'班'+'                     学科:';
    for I := 1 to 26 do
    begin
      titleInfo := titleInfo+' ';
    end;
    titleInfo := titleInfo+inttostr(Year)+'年'+inttostr(Month)+'月'+inttostr(Day)+'日';
    wordDoc.Paragraphs.Item(1).Range.InsertAfter(titleInfo);
  //从数据库中获取数据
    sql := 'select stuID          as 学号,' +
                  'stuName        as 姓名' +
              ' from tb_class_'+classID;
    ado.SelectInfo(ADOQuery, sql);
  //计算表格所需行数，并添加对应的行数
    if (ADOQuery.RecordCount mod 3) = 0 then
    begin
      dataRows := ADOQuery.RecordCount div 3;
    end else begin
      dataRows := (ADOQuery.RecordCount div 3) + 1;
    end;
    if (table.rows.Count-1) < dataRows then
    begin
      table.Select;
      wordApp.Selection.InsertRowsBelow(dataRows-(table.rows.Count-1));
    end;
  //向表格中插入数据
    i := 2;     //表头占一行，从第二行开始插入数据
    j := 1;
    while not ADOQuery.eof do
    begin
      //插入数据
      table.cell(i,j).Range.InsertAfter(ADOQuery.FieldByName('学号').AsString);
      table.cell(i,j+1).Range.InsertAfter(ADOQuery.FieldByName('姓名').AsString);
      //获取下一条记录
      ADOQuery.Next;
      i := (i mod (dataRows+1))+1;
      if i=1 then
      begin
        i := i+1;
        j := j+3;
      end;
    end;

    trans.Free;
    wordDoc.SaveAs(filename);
    if filePrint then
    begin
      wordApp.PrintOut;
    end;
    if visual then
    begin
      wordApp.visible := true;
    end else begin
      wordApp.Quit;
    end;
  end;

  procedure Output.stat_Excel_allCourse_school();
  var
    courseList : TStringList;
    course : Tb_course;
    i : integer;

    modelFile : string;
    path : string;
    fileName : string;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\登分表\全科登分表\Excel文件';
    fileName:= path + '\' +'所有年级学生成绩登分表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
  //获取该年级所有课程
    courseList := TStringList.Create;
    course := Tb_course.Create;
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //如有重复值则放弃
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('课程名称').AsString);
        ADOQuery.Next;
      end;
    end;

    for I := 0 to courseList.Count -1 do
    begin
      ExcelApp.ActiveSheet.Cells[1,i+3] := courseList[i];
    end;

    //将学生数据插入Excel表中，从第二行开始
    insterSchoolStuInfoToExcel(2);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;


  end;

  procedure Output.stat_Excel_allCourse_grade(grade : integer);
  var
    course : Tb_course;
    modelFile : string;
    path : string;
    fileName : string;

    i : integer;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\登分表\全科登分表\Excel文件';
    fileName:= path + '\' + inttostr(grade) + '年级学生成绩登分表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
  //获取这个年级所有课程
    course := Tb_course.Create;

    course.getCourseByGrade(ADOQuery, grade);
    i := 3;   //从第三列开始
    while not ADOQuery.eof do
    begin
      ExcelApp.ActiveSheet.Cells[1,i] := ADOQuery.FieldByName('课程名称').AsString;
      i := i+1;
      ADOQuery.Next;
    end;
    //将学生数据插入Excel表中，从第二行开始
    insterGradeStuInfoToExcel(2, grade);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.stat_Excel_allCourse_class(classID : string);
  var
   course : Tb_course;
    modelFile : string;
    path : string;
    fileName : string;

    trans : CTransform;
    grade : integer;

    i : integer;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\登分表\全科登分表\Excel文件';
    fileName:= path + '\' + classID + '班学生成绩登分表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
  //获取这个年级所有课程
    course := Tb_course.Create;
    trans := CTransform.Create;
    grade := trans.classIDToGrade(classID);

    course.getCourseByGrade(ADOQuery, grade);
    i := 3;   //从第三列开始
    while not ADOQuery.eof do
    begin
      ExcelApp.ActiveSheet.Cells[1,i] := ADOQuery.FieldByName('课程名称').AsString;
      i := i+1;
      ADOQuery.Next;
    end;
    //将学生数据插入Excel表中，从第二行开始
    insterClassStuInfoToExcel(2, classID);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;

  end;

  procedure Output.stat_Excel_aCourse_school();
  var
    modelFile : string;
    path : string;
    fileName : string;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\登分表\单科登分表\Excel文件';
    fileName:= path + '\' +'所有年级学生成绩登分表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //将学生数据插入Excel表中，从第二行开始
    insterSchoolStuInfoToExcel(2);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.stat_Excel_aCourse_grade(grade : integer);
  var
    modelFile : string;
    path : string;
    fileName : string;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\登分表\单科登分表\Excel文件';
    fileName:= path + '\' + inttostr(grade) + '年级学生成绩登分表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //将学生数据插入Excel表中，从第二行开始
    insterGradeStuInfoToExcel(2, grade);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.stat_Excel_aCourse_class(classID : string);
  var
    modelFile : string;
    path : string;
    fileName : string;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\登分表\单科登分表\Excel文件';
    fileName:= path + '\' + classID + '班学生成绩登分表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //将学生数据插入Excel表中，从第二行开始
    insterClassStuInfoToExcel(2, classID);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

{---------------------------------成绩表方法------------------------------------}
  procedure Output.score_word_allCourse_school();
  var
    sql : string;
  begin

  end;

  procedure Output.score_word_allCourse_grade();
  var
    sql : string;
  begin

  end;

  procedure Output.score_word_allCourse_class();
  var
    sql : string;
  begin

  end;

  procedure Output.score_word_aCourse_school(course : string; exam : string);
  var
    I : Integer;
    t_course : Tb_course;    //course表对象
  begin
    t_course := Tb_course.Create;
    for I := 1 to 9 do
    begin
      if t_course.isHasCourse(i, course) then     //判断i年级是否拥有课程Course
      begin
        score_word_aCourse_grade(i, course, exam);
      end;
    end;
  end;

  procedure Output.score_word_aCourse_grade(grade : integer; course : string; exam : string);
  var
    sql : string;
    ADOQuery1 : TADOQuery;
    classes : Tb_classes;
  begin
    ADOQuery1 := TADOQuery.Create(nil);
    classes := Tb_classes.Create;
  //获取这个年级所有的班级
    classes.selectByGrade(ADOQuery1, grade);
    while not ADOQuery1.Eof do
    begin
      score_word_aCourse_class(ADOQuery1.FieldByName('班号').AsString, course, exam);
      ADOQuery1.Next;
    end;
  end;

  procedure Output.score_word_aCourse_class(classID : string; course : string; exam : string);
  var
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    path : string;          //文件输出路径
    fileName : string;      //报表文件名
    i : integer;
    j : integer;
    dataRows : integer;
    sql : string;

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
    //复制模版并创建word对象打开模板
    modelFile := softPath + '\model\单科登分表.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\成绩表\单科成绩表\Word文件';
    fileName:= path + '\' + classID + '班学生成绩表'+exam+'_'+course+'.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //学生成绩统计表模板9列

    //加入表头信息
    trans := CTransform.Create;
    Present:= Now;
    DecodeDate(Present, Year, Month, Day);
    titleInfo := '班级:'+trans.classIDToText(classID)+'班'+'                     学科:'+course;
    for I := 1 to (26-Length(course)*2) do
    begin
      titleInfo := titleInfo+' ';
    end;
    titleInfo := titleInfo+inttostr(Year)+'年'+inttostr(Month)+'月'+inttostr(Day)+'日';
    wordDoc.Paragraphs.Item(1).Range.InsertAfter(titleInfo);


    //从数据库中获取数据
    sql := 'select  A.stuID          as 学号,' +
                    ' A.stuName        as 姓名,' +
                    ' B.'+exam+'       as 成绩'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
              ' where A.stuID = B.stuID';
    ado.SelectInfo(ADOQuery, sql);
    //计算表格所需行数，并添加对应的行数
    if (ADOQuery.RecordCount mod 3) = 0 then
    begin
      dataRows := ADOQuery.RecordCount div 3;
    end else begin
      dataRows := (ADOQuery.RecordCount div 3) + 1;
    end;
    if (table.rows.Count-1) < dataRows then
    begin
      table.Select;
      wordApp.Selection.InsertRowsBelow(dataRows-(table.rows.Count-1));
    end;

    i := 2;     //表头占一行，从第二行开始插入数据
    j := 1;
    while not ADOQuery.eof do
    begin
      //插入数据
      table.cell(i,j).Range.InsertAfter(ADOQuery.FieldByName('学号').AsString);
      table.cell(i,j+1).Range.InsertAfter(ADOQuery.FieldByName('姓名').AsString);
      table.cell(i,j+2).Range.InsertAfter(ADOQuery.FieldByName('成绩').AsString);
      //获取下一条记录
      ADOQuery.Next;
      i := (i mod (dataRows+1))+1;
      if i=1 then
      begin
        i := i+1;
        j := j+3;
      end;
    end;

    trans.Free;
    wordDoc.SaveAs(filename);
    if filePrint then
    begin
      wordApp.PrintOut;
    end;
    if visual then
    begin
      wordApp.visible := true;
    end else begin
      wordApp.Quit;
    end;

  end;

  procedure Output.score_Excel_allCourse_school(exam : string);
  var
    I : Integer;
  begin
    for I := 1 to 9 do
    begin
     score_Excel_allCourse_grade(i, exam);
    end;
  end;

  procedure Output.score_Excel_allCourse_grade(grade : integer; exam : string);
  var
    modelFile : string;
    path : string;
    fileName : string;
    sql : string;

    classIDList : TStringList;
    classes : Tb_classes;
    classID : string;

    trans : CTransform;
    courseList : TStringList;
    I : Integer;
    index : integer;
    indexTemp : integer;
  j: Integer;
  begin
    //创建新文件并打开文件
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\成绩表\全科成绩表\Excel文件';
    fileName:= path + '\' + inttostr(grade) + '年级学生成绩总表'+exam+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
    //获取该年级所有班级信息
    classIDList := TStringList.Create;
    classes := Tb_classes.Create;
    classes.selectByGrade(ADOQuery, grade);
    while not ADOQuery.eof do
    begin
      classIDList.Add(ADOQuery.FieldByName('班号').AsString);
      ADOQuery.Next;
    end;

    //获取该年级所有课程
    courseList := TStringList.Create;
    sql := 'SELECT courseName  from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
    while not ADOQuery.Eof do
    begin
      courseList.Add(ADOQuery.FieldByName('courseName').AsString);
      ADOQuery.next;
    end;

    index := 2;
    for I := 0 to classIDList.Count - 1 do
    begin
      classID := classIDList[i];

      //初始化前两列
      sql := 'select stuID, stuName from tb_class_'+classID+' order by stuID asc';
      ado.SelectInfo(ADOQuery, sql);
      indexTemp := index;
      while not ADOQuery.Eof do
      begin
        ExcelApp.ActiveSheet.Cells[index,1] := ADOQuery.FieldByName('stuID').AsString;
        ExcelApp.ActiveSheet.Cells[index,2] := ADOQuery.FieldByName('stuName').AsString;
        index := index+1;
        ADOQuery.Next;
      end;
      index := indexTemp;
      //插入分数
      for j := 0 to courseList.Count -1 do
      begin
        ExcelApp.ActiveSheet.Cells[1,j + 3] := '';
      end;
      index := addScoresExcel(classID,exam,courseList,index,ExcelApp);
    end;


    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;

    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.score_Excel_allCourse_class(classID : string; exam : string);
  var
    modelFile : string;
    path : string;
    fileName : string;
    sql : string;

    trans : CTransform;
    grade : integer;
    courseList : TStringList;
    I : Integer;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\成绩表\全科成绩表\Excel文件';
    fileName:= path + '\' + classID + '班学生成绩总表'+exam+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //获取该年级所有课程
    trans := CTransform.Create;
    grade := trans.classIDToGrade(classID);
    courseList := TStringList.Create;
    sql := 'SELECT courseName  from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
    while not ADOQuery.Eof do
    begin
      courseList.Add(ADOQuery.FieldByName('courseName').AsString);
      ADOQuery.next;
    end;

    //初始化前两列
    sql := 'select stuID, stuName from tb_class_'+classID+' order by stuID asc';
    ado.SelectInfo(ADOQuery, sql);
    i := 2;
    while not ADOQuery.Eof do
    begin
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('stuID').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('stuName').AsString;
      i := i+1;
      ADOQuery.Next;
    end;


    //插入分数
    ExcelApp.ActiveSheet.Cells[1,3] := '';
    addScoresExcel(classID,exam,courseList,2,ExcelApp);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;

    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.score_Excel_aCourse_school(course : string; exam : string);
  var
    I : Integer;
    t_course : Tb_course;    //course表对象
  begin
    t_course := Tb_course.Create;
    for I := 1 to 9 do
    begin
      if t_course.isHasCourse(i, course) then     //判断i年级是否拥有课程Course
      begin
        score_Excel_aCourse_grade(i, course, exam);
      end;
    end;
  end;

  procedure Output.score_Excel_aCourse_grade(grade : integer; course : string; exam : string);
  var
    modelFile : string;
    path : string;
    fileName : string;
    i : integer;
    trans : CTransform;
    classIDList : TStringList;
    classID : string;
    index : integer;
    classes : Tb_classes;

    sql : string;
  begin
  //创建新文件并打开文件
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\成绩表\单科成绩表\Excel文件';
    fileName:= path + '\' + inttostr(grade) + '年级学生成绩表'+exam+'_'+course+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
    //获取该年级所有班级信息
    classIDList := TStringList.Create;
    classes := Tb_classes.Create;
    classes.selectByGrade(ADOQuery, grade);
    while not ADOQuery.eof do
    begin
      classIDList.Add(ADOQuery.FieldByName('班号').AsString);
      ADOQuery.Next;
    end;
    //向Excel文件中插入学生信息
    index := 2;      //表头占一行，从第二行开始插入数据
    for I := 0 to classIDList.Count -1 do
    begin
      classID := classIDList[i];
      //从数据库中获取数据
      sql := 'select  A.stuID          as 学号,' +
                    ' A.stuName        as 姓名,' +
                    ' B.'+exam+'       as 成绩'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(grade)+'_'+course+' B'+
              ' where A.stuID = B.stuID';
      ado.SelectInfo(ADOQuery, sql);

      while not ADOQuery.eof do
      begin
        //插入数据
        ExcelApp.ActiveSheet.Cells[index,1] := ADOQuery.FieldByName('学号').AsString;
        ExcelApp.ActiveSheet.Cells[index,2] := ADOQuery.FieldByName('姓名').AsString;
        ExcelApp.ActiveSheet.Cells[index,3] := ADOQuery.FieldByName('成绩').AsString;
        //获取下一条记录
        ADOQuery.Next;
        index := index+1;
      end;
    end;

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.score_Excel_aCourse_class(classID : string; course : string; exam : string);
  var
    modelFile : string;
    path : string;
    fileName : string;
    i : integer;
    trans : CTransform;

    sql : string;
  begin
    modelFile := softPath + '\model\登分表.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\成绩表\单科成绩表\Excel文件';
    fileName:= path + '\' + classID + '班学生成绩表'+exam+'_'+course+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //从数据库中获取数据
    trans := CTransform.Create;
    sql := 'select  A.stuID          as 学号,' +
                  ' A.stuName        as 姓名,' +
                  ' B.'+exam+'       as 成绩'  +
            ' from tb_class_'+classID+ ' A,'+
            ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
            ' where A.stuID = B.stuID';
    trans.Free;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //表头占一行，从第二行开始插入数据
    while not ADOQuery.eof do
    begin
      //插入数据
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('学号').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('姓名').AsString;
      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('成绩').AsString;
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

{-------------------------------期末评估表方法----------------------------------}
  procedure Output.finalExam_word_allCourse_school(exam : string);
  var
    i : integer;
    courseList : TStringList;
    course : Tb_course;
  begin
    courseList := TStringList.Create;
    course := Tb_course.Create;
    //获取本学校所有课程
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //如有重复值则放弃
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('课程名称').AsString);
        ADOQuery.Next;
      end;
    end;
    for I := 0 to courseList.Count - 1 do
    begin
      finalExam_word_aCourse_school(courseList[i], exam);
    end;
  end;

  procedure Output.finalExam_word_aCourse_school(course : string; exam : string);
  var
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    path : string;          //
    fileName : string;      //报表文件名


    sql : string;
    i : integer;

  begin
    //复制模版并创建word对象打开模板
    modelFile := softPath + '\model\期末评估.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\期末评估\'+exam+'\Word文件';
    fileName:= path + '\' + course + '期末评估表.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 16);     //学生成绩统计表模板9列

    //修改标题
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.TypeText(course);

//    if comparestr('上学期期末',exam) = 0 then
//    begin
//      exam := '期末_上';
//      firstTermFianlExamDataToWord(exam, course, table);
//    end;
//
//    if comparestr('下学期期末',exam) = 0 then
//    begin
//      exam := '期末_下';
//      secondTermFianlExamDataToWord(exam, course, table);
//    end;
    exam := '期末';
    FianlExamDataToWord(exam, course, table);

    wordDoc.SaveAs(filename);

    if filePrint then
    begin
      wordApp.FilePrintPreview;
    end;

    if visual then
    begin
      wordApp.visible := true;
    end else begin
      wordApp.Quit;
    end;
  end;

  procedure Output.finalExam_Excel_allCourse_school(exam : string);
  var
    i : integer;
    courseList : TStringList;
    course : Tb_course;
  begin
    courseList := TStringList.Create;
    course := Tb_course.Create;
    //获取本学校所有课程
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //如有重复值则放弃
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('课程名称').AsString);
        ADOQuery.Next;
      end;
    end;
    for I := 0 to courseList.Count - 1 do
    begin
      finalExam_Excel_aCourse_school(courseList[i], exam);
    end;
  end;

  procedure Output.finalExam_Excel_aCourse_school(course : string; exam : string);
  var
    ExcelApp: OleVariant;
    modelFile : string;     //模板文件
    path : string;          //
    fileName : string;      //报表文件名


    sql : string;
    i : integer;

  begin
    //复制模版并创建word对象打开模板
    modelFile := softPath + '\model\期末评估.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\期末评估\'+exam+'\Excel文件';
    fileName:= path + '\' + course + '期末评估表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

//    if comparestr('上学期期末',exam) = 0 then
//    begin
//      exam := '期末_上';
//      FianlExamDataToExcel(exam, course, ExcelApp);
//    end;
//
//    if comparestr('下学期期末',exam) = 0 then
//    begin
//      exam := '期末_下';
//      secondTermFianlExamDataToExcel(exam, course, ExcelApp);
//    end;
    FianlExamDataToExcel(exam, course, ExcelApp);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.FianlExamDataToWord(exam : string; course : string; var table: OleVariant);
  var
    sql : string;
    I : integer;
    rowCount : integer;   //表格行数

    greatStuNum_a : string;
    greatStuNum_b : string;
    inferiorStuNum_a : string;
    inferiorStuNum_b : string;

    oldDatabaseName : string;
    connString : string;
    t_course : Tb_course;
    oldDatabaseHasCourse : boolean;
    className : string;
    trans : CTransform;

    ass : Tb_assessment;

    yearchange:boolean;
    moveyear:integer;

  begin
  //获取当前学年数据填入表中
    ass := Tb_assessment.Create;
    ass.createReport('期末', course);
    //从数据库中获取数据
    sql := 'select classID        as 班号,' +        //下学期班号
                  'testNum        as 实考_b,' +
                  'average        as 均分_b,' +
                  'standardDev    as 标准差_b,' +
                  'greatStuNum    as 优生_b,' +
                  'greatRatio     as 优生率_b,' +
                  'inferiorStuNum as 学困生_b,' +
                  'evaluate       as 异动_b,' +
                  'teacher        as 任课教师,' +    //下学期老师
                  'greatLine      as 优生线_b,' +
                  'inferiorLine   as 学困线_b' +
            ' from tb_report_期末_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //表头占两行，从第三行开始插入数据
    while not ADOQuery.eof do
    begin
      //表格中行数不够时增加行
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString +#13+
            '(' + ADOQuery.FieldByName('优生线_b').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString +#13+
            '(' + ADOQuery.FieldByName('学困线_b').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString;
      end;

      //插入数据
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('班号').AsString);
      table.cell(i,2).Range.InsertAfter('0');
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('实考_b').AsString);
      table.cell(i,4).Range.InsertAfter('0');
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('均分_b').AsString);
      table.cell(i,6).Range.InsertAfter('0');
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('标准差_b').AsString);
      table.cell(i,8).Range.InsertAfter('0');
      table.cell(i,9).Range.InsertAfter(greatStuNum_b);
      table.cell(i,10).Range.InsertAfter('0');
      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('优生率_b').AsString);
      table.cell(i,12).Range.InsertAfter('0');
      table.cell(i,13).Range.InsertAfter(inferiorStuNum_b);
      table.cell(i,14).Range.InsertAfter('');
      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('异动_b').AsString);
      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('任课教师').AsString);
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //表格行数


  //获取上一学年数据填入表中
    //检查上一学期数据库是否存在
    softPath := ExtractFileDir(ParamStr(0));      //获取软件本身所在路径，用于文件操作
    if comparestr(GlobalData.term,'上学期') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year-1)+'-'+inttostr(year)+'-下学期';
      yearchange := true;
      moveyear:=1;   //报表要求升年级后期末评估还是和自己班比较，所以上一学年一年级的数据和本学年二年级对应
    end else if comparestr(GlobalData.term,'下学期') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year)+'-'+inttostr(year+1)+'-上学期';
      yearchange := false;
      moveyear:=0;   //下学期和上学期期末对比，没有跨学年，所以不用再报表中移动一年
    end;


    if FileExists(oldDatabaseName+'.accdb') then
    begin
      oldDatabaseName := oldDatabaseName + '.accdb';
    end else if FileExists(oldDatabaseName+'.mdb') then begin
      oldDatabaseName := oldDatabaseName + '.mdb';
    end else begin
      exit;                    //没有找到上一学年数据库
    end;
    //保存现有数据库连接串
    connString := ADOOperate.connStr;
    //连接上一学期数据库
    ado.initConnStr(oldDatabaseName);
    if ado.testConn then
    begin
      if yearchange then
      begin
        //注意此时全局参数year没变
        GlobalData.year := GlobalData.year - 1;
      end;

      //判断当前学年数据库中是否有该课程
      t_course := Tb_course.Create;
      oldDatabaseHasCourse := false;
      for I := 1 to 9 do
      begin
        if t_course.isHasCourse(i, course) then
        begin
          oldDatabaseHasCourse := true;
          break;
        end;
      end;
      //如果旧数据库中有这门课程
      if oldDatabaseHasCourse then
      begin
        ass.createReport('期末', course);
        //从数据库中获取数据
        trans := CTransform.Create;
        for I := 3 to rowCount do
        begin
          className := table.cell(i,1).Range.Text;
          className := trans.numToText(trans.TextToNum(LeftStr(className, 1)) - moveyear)
              + RightStr(className, className.Length-1);
          className := LeftStr(className, className.Length-2);
          sql := 'select classID        as 班号,' +        //下学期班号
                      'testNum        as 实考_a,' +
                      'average        as 均分_a,' +
                      'standardDev    as 标准差_a,' +
                      'greatStuNum    as 优生_a,' +
                      'greatRatio     as 优生率_a,' +
                      'inferiorStuNum as 学困生_a,' +
                      'evaluate       as 异动_a,' +
                      'greatLine      as 优生线_a,' +
                      'inferiorLine   as 学困线_a ' +
                ' from tb_report_期末_' + course +
                ' where classID = ' + #39 + className + #39;
          ado.SelectInfo(ADOQuery, sql);
          //插入一条信息
          if ADOQuery.RecordCount <> 0 then
          begin
            if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
            begin
              greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('优生线_a').AsString + ')';
              inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('学困线_a').AsString + ')';
            end else begin
              greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString;
              inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString;
            end;

            //插入数据
            table.cell(i,2).Range.text := ADOQuery.FieldByName('实考_a').AsString;
            table.cell(i,4).Range.text := ADOQuery.FieldByName('均分_a').AsString;
            table.cell(i,6).Range.text := ADOQuery.FieldByName('标准差_a').AsString;
            table.cell(i,8).Range.text := greatStuNum_a;
            table.cell(i,10).Range.text := ADOQuery.FieldByName('优生率_a').AsString;
            table.cell(i,12).Range.text := inferiorStuNum_a;
            table.cell(i,14).Range.text := ADOQuery.FieldByName('异动_a').AsString;
          end;
        end;
      end;

      if yearchange then
      begin
        GlobalData.year := GlobalData.year + 1;
      end;
    end else begin
      showMessage('上一学年数据库连接失败');
    end;
    //恢复原有数据库连接
    ado.setConnStr(connString);
  end;

  procedure Output.secondTermFianlExamDataToWord(exam : string; course : string; var table: OleVariant);
//  var
//    sql : string;
//    I : integer;
//
//
//    greatStuNum_a : string;
//    greatStuNum_b : string;
//    inferiorStuNum_a : string;
//    inferiorStuNum_b : string;
//
//    ass : Tb_assessment;
  begin
//    ass := Tb_assessment.Create;
//    ass.createReport('期末_上', course);
//    ass.createReport('期末_下', course);
//    //从数据库中获取数据
//    sql := 'select B.classID        as 班号,' +        //下学期班号
//                  'A.testNum        as 实考_a,' +
//                  'B.testNum        as 实考_b,' +
//                  'A.average        as 均分_a,' +
//                  'B.average        as 均分_b,' +
//                  'A.standardDev    as 标准差_a,' +
//                  'B.standardDev    as 标准差_b,' +
//                  'A.greatStuNum    as 优生_a,' +
//                  'B.greatStuNum    as 优生_b,' +
//                  'A.greatRatio     as 优生率_a,' +
//                  'B.greatRatio     as 优生率_b,' +
//                  'A.inferiorStuNum as 学困生_a,' +
//                  'B.inferiorStuNum as 学困生_b,' +
//                  'A.evaluate       as 异动_a,' +
//                  'B.evaluate       as 异动_b,' +
//                  'B.teacher        as 任课教师,' +    //下学期老师
//                  'A.greatLine      as 优生线_a,' +
//                  'B.greatLine      as 优生线_b,' +
//                  'A.inferiorLine   as 学困线_a,' +
//                  'B.inferiorLine   as 学困线_b' +
//            ' from tb_report_期末_上_' + course+' A, tb_report_期末_下_' + course + ' B' +
//            ' where A.classID = B.classID';
//    ado.SelectInfo(ADOQuery, sql);
//
//    i := 3;     //表头占两行，从第三行开始插入数据
//    while not ADOQuery.eof do
//    begin
//      //表格中行数不够时增加行
//      if table.rows.Count < i then
//      begin
//        table.Select;
//        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
//      end;
//
//      if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
//      begin
//        greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('优生线_a').AsString + ')';
//        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('优生线_b').AsString + ')';
//        inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('学困线_a').AsString + ')';
//        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('学困线_b').AsString + ')';
//      end else begin
//        greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString;
//        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString;
//        inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString;
//        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString;
//      end;
//
//      //插入数据
//      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('班号').AsString);
//      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('实考_a').AsString);
//      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('实考_b').AsString);
//      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('均分_a').AsString);
//      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('均分_b').AsString);
//      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('标准差_a').AsString);
//      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('标准差_b').AsString);
//      table.cell(i,8).Range.InsertAfter(greatStuNum_a);
//      table.cell(i,9).Range.InsertAfter(greatStuNum_b);
//      table.cell(i,10).Range.InsertAfter(ADOQuery.FieldByName('优生率_a').AsString);
//      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('优生率_b').AsString);
//      table.cell(i,12).Range.InsertAfter(inferiorStuNum_a);
//      table.cell(i,13).Range.InsertAfter(inferiorStuNum_b);
//      table.cell(i,14).Range.InsertAfter(ADOQuery.FieldByName('异动_a').AsString);
//      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('异动_b').AsString);
//      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('任课教师').AsString);
//      //获取下一条记录
//      ADOQuery.Next;
//      i := i+1;
//    end;
  end;

  procedure Output.FianlExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
  var
    sql : string;
    I : integer;
    rowCount : integer;   //表格行数

    greatStuNum_a : string;
    greatStuNum_b : string;
    inferiorStuNum_a : string;
    inferiorStuNum_b : string;

    oldDatabaseName : string;
    connString : string;
    t_course : Tb_course;
    oldDatabaseHasCourse : boolean;
    className : string;
    trans : CTransform;

    ass : Tb_assessment;

    yearchange:boolean;
    moveyear:integer;
  begin
  //获取当前学年数据填入表中
    ass := Tb_assessment.Create;
    ass.createReport('期末', course);
    //从数据库中获取数据
    sql := 'select classID        as 班号,' +        //下学期班号
                  'testNum        as 实考_b,' +
                  'average        as 均分_b,' +
                  'standardDev    as 标准差_b,' +
                  'greatStuNum    as 优生_b,' +
                  'greatRatio     as 优生率_b,' +
                  'inferiorStuNum as 学困生_b,' +
                  'evaluate       as 异动_b,' +
                  'teacher        as 任课教师,' +    //下学期老师
                  'greatLine      as 优生线_b,' +
                  'inferiorLine   as 学困线_b' +
            ' from tb_report_期末_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //表头占两行，从第三行开始插入数据
    while not ADOQuery.eof do
    begin
      if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString +#13+
            '(' + ADOQuery.FieldByName('优生线_b').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString +#13+
            '(' + ADOQuery.FieldByName('学困线_b').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString;
      end;

      //插入数据
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('班号').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := '0';
      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('实考_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,4] := '0';
      ExcelApp.ActiveSheet.Cells[i,5] := ADOQuery.FieldByName('均分_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,6] := '0';
      ExcelApp.ActiveSheet.Cells[i,7] := ADOQuery.FieldByName('标准差_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,8] := '0';
      ExcelApp.ActiveSheet.Cells[i,9] := greatStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,10] := '0';
      ExcelApp.ActiveSheet.Cells[i,11] := ADOQuery.FieldByName('优生率_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,12] := '0';
      ExcelApp.ActiveSheet.Cells[i,13] := inferiorStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,14] := '';
      ExcelApp.ActiveSheet.Cells[i,15] := ADOQuery.FieldByName('异动_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,16] := ADOQuery.FieldByName('任课教师').AsString;
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //表格行数


  //获取上一学年数据填入表中
    //检查上一学期数据库是否存在
    softPath := ExtractFileDir(ParamStr(0));      //获取软件本身所在路径，用于文件操作
    if comparestr(GlobalData.term,'上学期') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year-1)+'-'+inttostr(year)+'-下学期';
      yearchange := true;
      moveyear:=1;   //报表要求升年级后期末评估还是和自己班比较，所以上一学年一年级的数据和本学年二年级对应
    end else if comparestr(GlobalData.term,'下学期') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year)+'-'+inttostr(year+1)+'-上学期';
      yearchange := false;
      moveyear:=0;   //下学期和上学期期末对比，没有跨学年，所以不用再报表中移动一年
    end;


    if FileExists(oldDatabaseName+'.accdb') then
    begin
      oldDatabaseName := oldDatabaseName + '.accdb';
    end else if FileExists(oldDatabaseName+'.mdb') then begin
      oldDatabaseName := oldDatabaseName + '.mdb';
    end else begin
      exit;                    //没有找到上一学年数据库
    end;
    //保存现有数据库连接串
    connString := ADOOperate.connStr;
    //连接上一学年数据库
    ado.initConnStr(oldDatabaseName);
    if ado.testConn then
    begin
      if yearchange then
      begin
        //注意此时全局参数year没变
        GlobalData.year := GlobalData.year - 1;
      end;

      //判断当前学年数据库中是否有该课程
      t_course := Tb_course.Create;
      oldDatabaseHasCourse := false;
      for I := 1 to 9 do
      begin
        if t_course.isHasCourse(i, course) then
        begin
          oldDatabaseHasCourse := true;
          break;
        end;
      end;
      //如果旧数据库中有这门课程
      if oldDatabaseHasCourse then
      begin
        ass.createReport('期末', course);
        //从数据库中获取数据
        trans := CTransform.Create;
        for I := 3 to rowCount do
        begin
          className := ExcelApp.ActiveSheet.Cells[i,1];
          className := trans.numToText(trans.TextToNum(LeftStr(className, 1)) - 1)
              + RightStr(className, className.Length-1);
          //className := LeftStr(className, className.Length-2);
          sql := 'select classID        as 班号,' +        //下学期班号
                      'testNum        as 实考_a,' +
                      'average        as 均分_a,' +
                      'standardDev    as 标准差_a,' +
                      'greatStuNum    as 优生_a,' +
                      'greatRatio     as 优生率_a,' +
                      'inferiorStuNum as 学困生_a,' +
                      'evaluate       as 异动_a,' +
                      'greatLine      as 优生线_a,' +
                      'inferiorLine   as 学困线_a ' +
                ' from tb_report_期末_' + course +
                ' where classID = ' + #39 + className + #39;
          ado.SelectInfo(ADOQuery, sql);
          //插入一条信息
          if ADOQuery.RecordCount <> 0 then
          begin
            if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
            begin
              greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('优生线_a').AsString + ')';
              inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('学困线_a').AsString + ')';
            end else begin
              greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString;
              inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString;
            end;

            //插入数据
            ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('实考_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,4] :=  ADOQuery.FieldByName('均分_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,6] := ADOQuery.FieldByName('标准差_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,8] := greatStuNum_a;
            ExcelApp.ActiveSheet.Cells[i,10] := ADOQuery.FieldByName('优生率_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,12] := inferiorStuNum_a;
            ExcelApp.ActiveSheet.Cells[i,14] := ADOQuery.FieldByName('异动_a').AsString;
          end;
        end;
      end;

      if yearchange then
      begin
        //注意此时全局参数year没变
        GlobalData.year := GlobalData.year + 1;
      end;
    end else begin
      showMessage('上一学年数据库连接失败');
    end;
    //恢复原有数据库连接
    ado.setConnStr(connString);

  end;

  procedure Output.secondTermFianlExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
//  var
//    sql : string;
//    I : integer;
//
//
//    greatStuNum_a : string;
//    greatStuNum_b : string;
//    inferiorStuNum_a : string;
//    inferiorStuNum_b : string;
//
//    ass : Tb_assessment;
  begin
//    ass := Tb_assessment.Create;
//    ass.createReport('期末_上', course);
//    ass.createReport('期末_下', course);
//    //从数据库中获取数据
//    sql := 'select B.classID        as 班号,' +        //下学期班号
//                  'A.testNum        as 实考_a,' +
//                  'B.testNum        as 实考_b,' +
//                  'A.average        as 均分_a,' +
//                  'B.average        as 均分_b,' +
//                  'A.standardDev    as 标准差_a,' +
//                  'B.standardDev    as 标准差_b,' +
//                  'A.greatStuNum    as 优生_a,' +
//                  'B.greatStuNum    as 优生_b,' +
//                  'A.greatRatio     as 优生率_a,' +
//                  'B.greatRatio     as 优生率_b,' +
//                  'A.inferiorStuNum as 学困生_a,' +
//                  'B.inferiorStuNum as 学困生_b,' +
//                  'A.evaluate       as 异动_a,' +
//                  'B.evaluate       as 异动_b,' +
//                  'B.teacher        as 任课教师,' +    //下学期老师
//                  'A.greatLine      as 优生线_a,' +
//                  'B.greatLine      as 优生线_b,' +
//                  'A.inferiorLine   as 学困线_a,' +
//                  'B.inferiorLine   as 学困线_b' +
//            ' from tb_report_期末_上_' + course+' A, tb_report_期末_下_' + course + ' B' +
//            ' where A.classID = B.classID';
//    ado.SelectInfo(ADOQuery, sql);
//
//    i := 3;     //表头占两行，从第三行开始插入数据
//    while not ADOQuery.eof do
//    begin
//      if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
//      begin
//        greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('优生线_a').AsString + ')';
//        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('优生线_b').AsString + ')';
//        inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('学困线_a').AsString + ')';
//        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('学困线_b').AsString + ')';
//      end else begin
//        greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString;
//        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString;
//        inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString;
//        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString;
//      end;
//
//      //插入数据
//      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('班号').AsString;
//      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('实考_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('实考_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,4] := ADOQuery.FieldByName('均分_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,5] := ADOQuery.FieldByName('均分_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,6] := ADOQuery.FieldByName('标准差_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,7] := ADOQuery.FieldByName('标准差_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,8] := greatStuNum_a;
//      ExcelApp.ActiveSheet.Cells[i,9] := greatStuNum_b;
//      ExcelApp.ActiveSheet.Cells[i,10] := ADOQuery.FieldByName('优生率_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,11] := ADOQuery.FieldByName('优生率_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,12] := inferiorStuNum_a;
//      ExcelApp.ActiveSheet.Cells[i,13] := inferiorStuNum_b;
//      ExcelApp.ActiveSheet.Cells[i,14] := ADOQuery.FieldByName('异动_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,15] := ADOQuery.FieldByName('异动_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,16] := ADOQuery.FieldByName('任课教师').AsString;
//      //获取下一条记录
//      ADOQuery.Next;
//      i := i+1;
//    end;
  end;

{-------------------------------考试评估表方法----------------------------------}
  procedure Output.exam_word_allCourse_school( exam : string);
  var
    i : integer;
    courseList : TStringList;
    course : Tb_course;
  begin
    courseList := TStringList.Create;
    course := Tb_course.Create;
    //获取本学校所有课程
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //如有重复值则放弃
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('课程名称').AsString);
        ADOQuery.Next;
      end;
    end;
    for I := 0 to courseList.Count - 1 do
    begin
      exam_word_aCourse_school(courseList[i], exam);
    end;
  end;

  procedure Output.exam_word_aCourse_school(course : string; exam : string);
   var
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    path : string;          //
    fileName : string;      //报表文件名


    sql : string;
    i : integer;

  begin
    //复制模版并创建word对象打开模板
    modelFile := softPath + '\model\考试评估.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\考试评估\'+exam+'\Word文件';
    fileName:= path + '\' +exam+'-'+ course + '-考试评估表.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //学生成绩统计表模板9列

    //修改标题
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.TypeText(course);


    ExamDataToWord(exam, course, table);

    wordDoc.SaveAs(filename);

    if filePrint then
    begin
      wordApp.FilePrintPreview;
    end;

    if visual then
    begin
      wordApp.visible := true;
    end else begin
      wordApp.Quit;
    end;
  end;

  procedure Output.exam_Excel_allCourse_school( exam : string);
  var
    i : integer;
    courseList : TStringList;
    course : Tb_course;
  begin
    courseList := TStringList.Create;
    course := Tb_course.Create;
    //获取本学校所有课程
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //如有重复值则放弃
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('课程名称').AsString);
        ADOQuery.Next;
      end;
    end;
    for I := 0 to courseList.Count - 1 do
    begin
      exam_Excel_aCourse_school(courseList[i], exam);
    end;
  end;

  procedure Output.exam_Excel_aCourse_school(course : string; exam : string);
   var
    ExcelApp: OleVariant;
    modelFile : string;     //模板文件
    path : string;          //
    fileName : string;      //报表文件名


    sql : string;
    i : integer;

  begin
    //复制模版并创建word对象打开模板
    modelFile := softPath + '\model\考试评估.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\考试评估\'+exam+'\Excel文件';
    fileName:= path + '\'+exam+'-' + course + '-考试评估表.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    ExamDataToExcel(exam, course, ExcelApp);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;


  procedure Output.ExamDataToWord(exam : string; course : string; var table: OleVariant);
  var
    sql : string;
    I : integer;
    rowCount : integer;   //表格行数

    greatStuNum_a : string;
    greatStuNum_b : string;
    inferiorStuNum_a : string;
    inferiorStuNum_b : string;

    oldDatabaseName : string;
    connString : string;
    t_course : Tb_course;
    oldDatabaseHasCourse : boolean;
    className : string;
    trans : CTransform;

    ass : Tb_assessment;
  begin
  //获取当前学年数据填入表中
    ass := Tb_assessment.Create;
    ass.createReport(exam, course);
    //从数据库中获取数据
    sql := 'select classID        as 班号,' +
                  'testNum        as 实考,' +
                  'average        as 均分,' +
                  'standardDev    as 标准差,' +
                  'greatStuNum    as 优生,' +
                  'greatRatio     as 优生率,' +
                  'inferiorStuNum as 学困生,' +
                  'evaluate       as 异动,' +
                  'teacher        as 任课教师,' +
                  'greatLine      as 优生线,' +
                  'inferiorLine   as 学困线' +
            ' from tb_report_'+exam+'_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //表头占两行，从第三行开始插入数据
    while not ADOQuery.eof do
    begin
      //表格中行数不够时增加行
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('优生').AsString +#13+
            '(' + ADOQuery.FieldByName('优生线').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('学困生').AsString +#13+
            '(' + ADOQuery.FieldByName('学困线').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('优生').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('学困生').AsString;
      end;

      //插入数据
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('班号').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('实考').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('均分').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('标准差').AsString);
      table.cell(i,5).Range.InsertAfter(greatStuNum_b);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('优生率').AsString);
      table.cell(i,7).Range.InsertAfter(inferiorStuNum_b);
      table.cell(i,8).Range.InsertAfter(ADOQuery.FieldByName('异动').AsString);
      table.cell(i,9).Range.InsertAfter(ADOQuery.FieldByName('任课教师').AsString);
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //表格行数
  end;


  procedure Output.ExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
  var
    sql : string;
    I : integer;
    rowCount : integer;   //表格行数

    greatStuNum_a : string;
    greatStuNum_b : string;
    inferiorStuNum_a : string;
    inferiorStuNum_b : string;

    oldDatabaseName : string;
    connString : string;
    t_course : Tb_course;
    oldDatabaseHasCourse : boolean;
    className : string;
    trans : CTransform;

    ass : Tb_assessment;

    yearchange:boolean;
  begin
  //获取当前学年数据填入表中
    ass := Tb_assessment.Create;
    ass.createReport(exam, course);
    //从数据库中获取数据
    sql := 'select classID        as 班号,' +
                  'testNum        as 实考,' +
                  'average        as 均分,' +
                  'standardDev    as 标准差,' +
                  'greatStuNum    as 优生,' +
                  'greatRatio     as 优生率,' +
                  'inferiorStuNum as 学困生,' +
                  'evaluate       as 异动,' +
                  'teacher        as 任课教师,' +
                  'greatLine      as 优生线,' +
                  'inferiorLine   as 学困线' +
            ' from tb_report_'+exam+'_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //表头占两行，从第二行开始插入数据
    while not ADOQuery.eof do
    begin
      if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('优生').AsString +#13+
            '(' + ADOQuery.FieldByName('优生线').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('学困生').AsString +#13+
            '(' + ADOQuery.FieldByName('学困线').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('优生').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('学困生').AsString;
      end;

      //插入数据
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('班号').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('实考').AsString;
      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('均分').AsString;
      ExcelApp.ActiveSheet.Cells[i,4] := ADOQuery.FieldByName('标准差').AsString;
      ExcelApp.ActiveSheet.Cells[i,5] := greatStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,6] := ADOQuery.FieldByName('优生率').AsString;
      ExcelApp.ActiveSheet.Cells[i,7] := inferiorStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,8] := ADOQuery.FieldByName('异动').AsString;
      ExcelApp.ActiveSheet.Cells[i,9] := ADOQuery.FieldByName('任课教师').AsString;
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //表格行数
  end;

{***************************************Word**************************************}
  procedure Output.finalreport(course : string);
  var
//    wordApp: OleVariant;
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    fileName : string;      //报表文件名

    greatStuNum_a : string;
    greatStuNum_b : string;
    inferiorStuNum_a : string;
    inferiorStuNum_b : string;

    sql : string;
    i : integer;
  begin
    modelFile := softPath + '\model\期末评估.docx';
    fileName:=softPath + '\report\期末评估\期末评估_'
        +course+inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1)+'.docx';
//    createWord(modelFile, fileName, wordApp, wordDoc, table, 16);     //期末评估模板16列

    //修改标题
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.TypeText(course);

    //从数据库中获取数据
    sql := 'select A.classID        as 班号,' +
                  'A.testNum        as 实考_a,' +
                  'B.testNum        as 实考_b,' +
                  'A.average        as 均分_a,' +
                  'B.average        as 均分_b,' +
                  'A.standardDev    as 标准差_a,' +
                  'B.standardDev    as 标准差_b,' +
                  'A.greatStuNum    as 优生_a,' +
                  'B.greatStuNum    as 优生_b,' +
                  'A.greatRatio     as 优生率_a,' +
                  'B.greatRatio     as 优生率_b,' +
                  'A.inferiorStuNum as 学困生_a,' +
                  'B.inferiorStuNum as 学困生_b,' +
                  'A.evaluate       as 异动_a,' +
                  'B.evaluate       as 异动_b,' +
                  'A.teacher        as 任课教师,' +
                  'A.greatLine      as 优生线_a,' +
                  'B.greatLine      as 优生线_b,' +
                  'A.inferiorLine   as 学困线_a,' +
                  'B.inferiorLine   as 学困线_b' +
            ' from tb_report_期末_上_'+course+' A,tb_report_期末_下_'+course+' B' +
            ' where A.classID = B.classID';
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //表头占两行，从第三行开始插入数据
    while not ADOQuery.eof do
    begin
      //表格中行数不够时增加行
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      if comparestr('年级',RightStr(ADOQuery.FieldByName('班号').AsString,2)) = 0 then
      begin
        greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString +#13+
            '(' + ADOQuery.FieldByName('优生线_a').AsString + ')';
        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString +#13+
            '(' + ADOQuery.FieldByName('优生线_b').AsString + ')';
        inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString +#13+
            '(' + ADOQuery.FieldByName('学困线_a').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString +#13+
            '(' + ADOQuery.FieldByName('学困线_b').AsString + ')';
      end else begin
        greatStuNum_a := ADOQuery.FieldByName('优生_a').AsString;
        greatStuNum_b := ADOQuery.FieldByName('优生_b').AsString;
        inferiorStuNum_a := ADOQuery.FieldByName('学困生_a').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('学困生_b').AsString;
      end;


      //插入数据
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('班号').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('实考_a').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('实考_b').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('均分_a').AsString);
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('均分_b').AsString);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('标准差_a').AsString);
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('标准差_b').AsString);
      table.cell(i,8).Range.InsertAfter(greatStuNum_a);
      table.cell(i,9).Range.InsertAfter(greatStuNum_b);
      table.cell(i,10).Range.InsertAfter(ADOQuery.FieldByName('优生率_a').AsString);
      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('优生率_b').AsString);
      table.cell(i,12).Range.InsertAfter(inferiorStuNum_a);
      table.cell(i,13).Range.InsertAfter(inferiorStuNum_b);
      table.cell(i,14).Range.InsertAfter(ADOQuery.FieldByName('异动_a').AsString);
      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('异动_b').AsString);
      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('任课教师').AsString);
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;

    wordDoc.SaveAs(filename);

    if filePrint then
    begin
      wordApp.FilePrintPreview;
    end;

    if visual then
    begin
      wordApp.visible := true;
    end else begin
      wordApp.Quit;
    end;

  end;

  procedure Output.examReport(examName, course : string);
  var
//    wordApp: OleVariant;
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    fileName : string;      //报表文件名
    i : integer;
    j : integer;
    dataRows : integer;

    sql : string;
  begin
    modelFile := softPath + '\model\考试评估.docx';
    fileName:= softPath + '\report\考试评估\考试评估_'+examName+'.docx';
//    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //考试评估模板9列

    //修改标题
    for I := 1 to 6 do
    begin
      wordApp.Selection.delete;
    end;
    wordApp.Selection.TypeText(course+'学科'+examName);

    //从数据库中获取数据
    sql := 'select classID        as 班号,' +
                  'testNum        as 实考,' +
                  'average        as 均分,' +
                  'standardDev    as 标准差,' +
                  'greatStuNum    as 优生,' +
                  'greatRatio     as 优生率,' +
                  'inferiorStuNum as 学困生,' +
                  'evaluate       as 异动,' +
                  'teacher        as 任课教师' +
            ' from tb_report_'+examName+'_'+course;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //表头占一行，从第二行开始插入数据
    while not ADOQuery.eof do
    begin
      //表格中行数不够时增加行
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      //插入数据
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('班号').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('实考').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('均分').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('标准差').AsString);
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('优生').AsString);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('优生率').AsString);
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('学困生').AsString);
      table.cell(i,8).Range.InsertAfter(ADOQuery.FieldByName('异动').AsString);
      table.cell(i,9).Range.InsertAfter(ADOQuery.FieldByName('任课教师').AsString);
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;

    wordDoc.SaveAs(filename);
    if filePrint then
    begin
      wordApp.PrintOut;
    end;
    if visual then
    begin
      wordApp.visible := true;
    end else begin
      wordApp.Quit;
    end;
  end;

  procedure Output.roster(classID : string;course,exam : string);
  var
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    path : string;          //文件输出路径
    fileName : string;      //报表文件名
    i : integer;
    j : integer;
    dataRows : integer;
    sql : string;

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
    modelFile := softPath + '\model\学生成绩统计表.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        '学年-'+GlobalData.term+'\成绩统计表';
    fileName:= path + '\' + classID + '班学生成绩统计表.docx';
    if Comparestr('',course) <> 0 then
    begin
      fileName := softPath + '\report\'+exam+'_'+course+'_'+classID+'.docx';
    end;

//    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //学生成绩统计表模板9列

    //加入表头信息
    trans := CTransform.Create;
    Present:= Now;
    DecodeDate(Present, Year, Month, Day);
    titleInfo := '班级:'+trans.classIDToText(classID)+'班'+'                     学科:'+course;

    for I := 1 to (26-Length(course)*2) do
    begin
      titleInfo := titleInfo+' ';
    end;

    titleInfo := titleInfo+inttostr(Year)+'年'+inttostr(Month)+'月'+inttostr(Day)+'日';

    wordDoc.Paragraphs.Item(1).Range.InsertAfter(titleInfo);


    //从数据库中获取数据
    if Comparestr('',course) = 0 then
    begin
      sql := 'select stuID          as 学号,' +
                    'stuName        as 姓名' +
              ' from tb_class_'+classID;
    end else begin
      sql := 'select  A.stuID          as 学号,' +
                    ' A.stuName        as 姓名,' +
                    ' B.'+exam+'       as 成绩'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
              ' where A.stuID = B.stuID';
    end;
    ado.SelectInfo(ADOQuery, sql);
    //计算表格所需行数，并添加对应的行数
    if (ADOQuery.RecordCount mod 3) = 0 then
    begin
      dataRows := ADOQuery.RecordCount div 3;
    end else begin
      dataRows := (ADOQuery.RecordCount div 3) + 1;
    end;
    if (table.rows.Count-1) < dataRows then
    begin
      table.Select;
      wordApp.Selection.InsertRowsBelow(dataRows-(table.rows.Count-1));
    end;

    i := 2;     //表头占一行，从第二行开始插入数据
    j := 1;
    while not ADOQuery.eof do
    begin
      //插入数据
      table.cell(i,j).Range.InsertAfter(ADOQuery.FieldByName('学号').AsString);
      table.cell(i,j+1).Range.InsertAfter(ADOQuery.FieldByName('姓名').AsString);
      if Comparestr('',course) <> 0 then
      begin
        table.cell(i,j+2).Range.InsertAfter(ADOQuery.FieldByName('成绩').AsString);
      end;
      //获取下一条记录
      ADOQuery.Next;
      i := (i mod (dataRows+1))+1;
      if i=1 then
      begin
        i := i+1;
        j := j+3;
      end;
    end;

    trans.Free;
    wordDoc.SaveAs(filename);
    if filePrint then
    begin
      wordApp.PrintOut;
    end;
    if visual then
    begin
      wordApp.visible := true;
    end else begin
      wordApp.Quit;
    end;
  end;

  procedure Output.createWord(modelFile, path,fileName : string;
      var wordApp, wordDoc, table : OleVariant; Columns : integer);
  begin
  //判断文件路径是否存在,若不存在则先创建相关文件夹
    if not DirectoryExists(path) then //判断目录是否存在
    try
      begin
        ForceDirectories(path);   //创建目录
      end;
    except
      raise Exception.Create('无法建立路径： '+path);
    end;
  //复制模板文件
    CopyFile(PChar(modelFile), PChar(filename), False);      //false表示替换同名旧文件
  //创建word对象，并打开文件
    try
      wordApp := CreateOleObject('Word.Application');
      wordApp.Visible := false;
    except
      ShowMessage('创建word对象失败！请确认您使用的电脑是否安装有word');
      Exit;
    end;
    wordDoc := WordApp.Documents.Open(filename,false,false,false);
    //检查模板是否符合规范
    try
      table := WordDoc.Tables.Item(1);
      if table.Columns.Count <> Columns then
      begin
        ShowMessage('模板中表格不符合规范，应该有'+inttostr(Columns)+'列！');
        wordApp.Quit;
        Exit;
      end;
    except
      ShowMessage('模板中表格丢失！');
      wordApp.Quit;
      Exit;
    end;
  end;

{***************************************Excel*************************************}
  procedure Output.rosterExcel(classID : string;course,exam : string);
  var
    modelFile : string;
    fileName : string;
    i : integer;
    trans : CTransform;

    sql : string;
  begin
    modelFile := softPath + '\model\学生成绩统计表.xlsx';
    fileName := softPath + '\report\学生成绩统计表_'+classID+'.xlsx';
    if Comparestr('',course) <> 0 then
    begin
      fileName := softPath + '\report\'+exam+'_'+course+'_'+classID+'.xlsx';
    end;

//    createExcel(modelFile, fileName, ExcelApp);

    //从数据库中获取数据
    if Comparestr('',course) = 0 then
    begin
      sql := 'select stuID          as 学号,' +
                    'stuName        as 姓名' +
              ' from tb_class_'+classID;
    end else begin
      trans := CTransform.Create;
      sql := 'select  A.stuID          as 学号,' +
                    ' A.stuName        as 姓名,' +
                    ' B.'+exam+'       as 成绩'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
              ' where A.stuID = B.stuID';
      trans.Free;
    end;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //表头占一行，从第二行开始插入数据
    while not ADOQuery.eof do
    begin
      //插入数据
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('学号').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('姓名').AsString;
      if Comparestr('',course) <> 0 then
      begin
        ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('成绩').AsString;
      end;


      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;
    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;
  end;

  procedure Output.allScoresExcel(classID : string;exam : string);
  var
//    ExcelApp: OleVariant;
    modelFile : string;
    fileName : string;
    sql : string;

    trans : CTransform;
    grade : integer;
    courseList : TStringList;
    I : Integer;
  begin
    modelFile := softPath + '\model\成绩总表.xlsx';
    fileName := softPath + '\report\成绩总表_'+exam+'_'+classID+'.xlsx';
//    createExcel(modelFile, fileName, ExcelApp);

    //获取该年级所有课程
    trans := CTransform.Create;
    grade := trans.classIDToGrade(classID);
    courseList := TStringList.Create;
    sql := 'SELECT courseName  from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
    while not ADOQuery.Eof do
    begin
      courseList.Add(ADOQuery.FieldByName('courseName').AsString);
      ADOQuery.next;
    end;

    //初始化前两列
    sql := 'select stuID, stuName from tb_class_'+classID+' order by stuID asc';
    ado.SelectInfo(ADOQuery, sql);
    i := 2;
    while not ADOQuery.Eof do
    begin
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('stuID').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('stuName').AsString;
      i := i+1;
      ADOQuery.Next;
    end;

    //插入分数
    addScoresExcel(classID,exam,courseList,2,ExcelApp);

    ExcelApp.ActiveWorkBook.Save;
    if filePrint then
    begin
      ExcelApp.ActiveSheet.PrintOut;
    end;

    if visual then
    begin
      ExcelApp.visible := true;
    end else begin
      ExcelApp.Quit;
    end;

  end;

  procedure Output.createExcel(modelFile, path, fileName : string;var ExcelApp : OleVariant);
  begin
  //判断文件路径是否存在,若不存在则先创建相关文件夹
    if not DirectoryExists(path) then //判断目录是否存在
    try
      begin
        ForceDirectories(path);   //创建目录
      end;
    except
      raise Exception.Create('无法建立路径： '+path);
    end;
    CopyFile(PChar(modelFile), PChar(filename), False);      //false表示替换同名旧文件

    try
      ExcelApp := CreateOleObject('Excel.Application');
      ExcelApp.Visible := false;
    except
      ShowMessage('创建excel对象失败！请确认您使用的电脑是否安装有excel');
      Exit;
    end;
    ExcelApp.WorkBooks.Open(fileName);

    ExcelApp.ActiveSheet.Rows[1].Font.Name := '宋体';
    ExcelApp.ActiveSheet.Rows[1].Font.Bold := True;
  end;

  function Output.addScoresExcel(classID, exam : string; var courseList : TstringList;
    rowIndex : integer; var ExcelApp : OleVariant):integer;
  var
    sql : string;
    i : integer;
    j : integer;
    columnIndex : integer;
    trans : CTransform;
    grade : integer;
  begin
    trans := CTransform.Create;
    grade := trans.classIDToGrade(classID);

    columnIndex := 1;
    while Comparestr('',ExcelApp.ActiveSheet.Cells[1,columnIndex]) <> 0 do
    begin
      columnIndex := columnIndex + 1;
    end;


    for I := 0 to courseList.Count-1 do
    begin
      ExcelApp.ActiveSheet.Cells[1,columnIndex] := courseList[i];  //设表头

      sql := 'select stuID, '+exam+' as 成绩 from tb_scores_'+inttostr(grade)+'_'+courseList[i]+
          ' where classID = '+#39+classID+#39+
          ' order by stuID asc';
      ado.SelectInfo(ADOQuery, sql);

      while not ADOQuery.Eof do
      begin
        j := rowIndex;
        while Comparestr('',ExcelApp.ActiveSheet.Cells[j,1]) <> 0 do
        begin
          if Comparestr(ADOQuery.FieldByName('stuID').AsString, ExcelApp.ActiveSheet.Cells[j,1]) =0 then
          begin
             ExcelApp.ActiveSheet.Cells[j,columnIndex] := ADOQuery.FieldByName('成绩').Asfloat;
             break;
          end;
          j := j + 1;
        end;
        ADOQuery.Next;
      end;
      columnIndex := columnIndex + 1;
    end;
    result := j;
  end;

  function Output.insterClassStuInfoToExcel(index : integer; classID : string):integer;
  var
    sql : string;
  begin
    //从数据库中获取数据
    sql := 'select stuID      as 学号,' +
              ' stuName       as 姓名' +
              ' from tb_class_'+classID;
    ado.SelectInfo(ADOQuery, sql);

    while not ADOQuery.eof do
    begin
      //插入数据
      ExcelApp.ActiveSheet.Cells[index,1] := ADOQuery.FieldByName('学号').AsString;
      ExcelApp.ActiveSheet.Cells[index,2] := ADOQuery.FieldByName('姓名').AsString;
      //获取下一条记录
      ADOQuery.Next;
      index := index+1;
    end;
    result := index;
  end;

  function Output.insterGradeStuInfoToExcel(index : integer; grade : integer):integer;
  var
    classes : Tb_classes;
    ADOQuery1 : TADOQuery;
  begin
    classes := Tb_classes.Create;
    ADOQuery1 := TADOQuery.Create(nil);

    classes.selectByGrade(ADOQuery1, grade);
    while not ADOQuery1.EOF do
    begin
      index := insterClassStuInfoToExcel(index,ADOQuery1.FieldByName('班号').AsString);
      ADOQuery1.Next;
    end;
    result := index;
  end;

  function Output.insterSchoolStuInfoToExcel(index : integer):integer;
  var
    I : Integer;
  begin
    for I := 1 to 9 do
    begin
      index := insterGradeStuInfoToExcel(index, i);
    end;
  end;
end.
