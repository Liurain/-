unit fileOutput;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Vcl.OleServer, Word2000, Comobj,DataTransform,StrUtils ;

type
    Output = class
    Constructor Create;
  private
    ado : TAdoOperate;        //数据库操作对象
    ADOQuery : TADOQuery;
    softPath : string;
    visual : boolean; //是否预览



    procedure createWord(modelFile, fileName : string;
        var wordApp, wordDoc, table : OleVariant; Columns : integer);
    procedure createExcel(modelFile, fileName : string;var ExcelApp : OleVariant);
    procedure addScoresExcel(classID, exam : string; var courseList : TstringList;var ExcelApp : OleVariant);
  public
    procedure finalReport(course : string);
    procedure examReport(examName, course : string);
    procedure roster(classID : string;course,exam : string);
    procedure rosterExcel(classID : string;course,exam : string);
    procedure allScoresExcel(classID : string;exam : string);

    procedure isVisual(flag : boolean);
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
      wordApp.Quit;           //关闭上一次预览未关闭的窗体并释放文件访问
      ExcelApp.Quit;
    except
    end;

  end;

  procedure Output.isVisual(flag : boolean);
  begin
    visual := flag;
  end;
{***************************************Word**************************************}
  procedure Output.finalreport(course : string);
  var
//    wordApp: OleVariant;
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //模板文件
    fileName : string;      //报表文件名



    sql : string;
    i : integer;
  begin
    modelFile := softPath + '\model\期末评估.docx';
    fileName:=softPath + '\report\期末评估\期末评估.docx';
    createWord(modelFile, fileName, wordApp, wordDoc, table, 16);     //期末评估模板16列

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
                  'A.teacher        as 任课教师' +
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



      //插入数据
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('班号').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('实考_a').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('实考_b').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('均分_a').AsString);
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('均分_b').AsString);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('标准差_a').AsString);
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('标准差_b').AsString);
      table.cell(i,8).Range.InsertAfter(ADOQuery.FieldByName('优生_a').AsString);
      table.cell(i,9).Range.InsertAfter(ADOQuery.FieldByName('优生_b').AsString);
      table.cell(i,10).Range.InsertAfter(ADOQuery.FieldByName('优生率_a').AsString);
      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('优生率_b').AsString);
      table.cell(i,12).Range.InsertAfter(ADOQuery.FieldByName('学困生_a').AsString);
      table.cell(i,13).Range.InsertAfter(ADOQuery.FieldByName('学困生_b').AsString);
      table.cell(i,14).Range.InsertAfter(ADOQuery.FieldByName('异动_a').AsString);
      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('异动_b').AsString);
      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('任课教师').AsString);
      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;

    wordDoc.SaveAs(filename);
    if not visual then wordApp.Quit;
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
    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //考试评估模板9列

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
    if not visual then wordApp.Quit;
  end;

  procedure Output.roster(classID : string;course,exam : string);
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

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
    modelFile := softPath + '\model\学生成绩统计表.docx';
    fileName:= softPath + '\report\成绩统计表\学生成绩统计表'+classID+'.docx';
    if Comparestr('',course) <> 0 then
    begin
      fileName := softPath + '\report\'+exam+'_'+course+'_'+classID+'.docx';
    end;

    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //学生成绩统计表模板9列

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
    if not visual then wordApp.Quit;

  end;

  procedure Output.createWord(modelFile, fileName : string;
      var wordApp, wordDoc, table : OleVariant; Columns : integer);
  begin
    CopyFile(PChar(modelFile), PChar(filename), False);      //false表示替换同名旧文件

//    try
//      wordApp := GetActiveOleObject('Word.Application');
//    except
      try
        wordApp := CreateOleObject('Word.Application');
        wordApp.Visible := Visual;
      except
        ShowMessage('创建word对象失败！请确认您使用的电脑是否安装有word');
        Exit;
      end;
//    end;
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
//    ExcelApp: OleVariant;
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

    createExcel(modelFile, fileName, ExcelApp);

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
      ExcelApp.WorkSheets[1].Cells[i,1].Value := ADOQuery.FieldByName('学号').AsString;
      ExcelApp.WorkSheets[1].Cells[i,2].Value := ADOQuery.FieldByName('姓名').AsString;
      if Comparestr('',course) <> 0 then
      begin
        ExcelApp.WorkSheets[1].Cells[i,3].Value := ADOQuery.FieldByName('成绩').AsString;
      end;


      //获取下一条记录
      ADOQuery.Next;
      i := i+1;
    end;

    ExcelApp.ActiveWorkBook.Save;
    if not visual then ExcelApp.Quit;
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
    createExcel(modelFile, fileName, ExcelApp);

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
      ExcelApp.WorkSheets[1].Cells[i,1].Value := ADOQuery.FieldByName('stuID').AsString;
      ExcelApp.WorkSheets[1].Cells[i,2].Value := ADOQuery.FieldByName('stuName').AsString;
      i := i+1;
      ADOQuery.Next;
    end;

    //插入分数
    addScoresExcel(classID,exam,courseList,ExcelApp);

    ExcelApp.ActiveWorkBook.Save;
    if not visual then ExcelApp.Quit;
  end;

  procedure Output.createExcel(modelFile, fileName : string;var ExcelApp : OleVariant);
  begin
    CopyFile(PChar(modelFile), PChar(filename), False);      //false表示替换同名旧文件

//    try
//      ExcelApp := GetActiveOleObject('Excel.Application');
//    except
      try
        ExcelApp := CreateOleObject('Excel.Application');
        ExcelApp.Visible := Visual;
      except
        ShowMessage('创建excel对象失败！请确认您使用的电脑是否安装有excel');
        Exit;
      end;
//    end;
    ExcelApp.WorkBooks.Open(fileName);

    ExcelApp.ActiveSheet.Rows[1].Font.Name := '宋体';
    ExcelApp.ActiveSheet.Rows[1].Font.Bold := True;
  end;

  procedure Output.addScoresExcel(classID, exam : string; var courseList : TstringList; var ExcelApp : OleVariant);
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
    while Comparestr('',ExcelApp.WorkSheets[1].Cells[1,columnIndex].Value) <> 0 do
    begin
      columnIndex := columnIndex + 1;
    end;


    for I := 0 to courseList.Count-1 do
    begin
      ExcelApp.WorkSheets[1].Cells[1,columnIndex].Value := courseList[i];  //设表头

      sql := 'select stuID, '+exam+' as 成绩 from tb_scores_'+inttostr(grade)+'_'+courseList[i]+
          ' where classID = '+#39+classID+#39+
          ' order by stuID asc';
      ado.SelectInfo(ADOQuery, sql);

      while not ADOQuery.Eof do
      begin
        j := 2;
        while Comparestr('',ExcelApp.WorkSheets[1].Cells[j,1].Value) <> 0 do
        begin
          if Comparestr(ADOQuery.FieldByName('stuID').AsString, ExcelApp.WorkSheets[1].Cells[j,1].Value) =0 then
          begin
             ExcelApp.WorkSheets[1].Cells[j,columnIndex].Value := ADOQuery.FieldByName('成绩').Asfloat;
             break;
          end;
          j := j + 1;
        end;
        ADOQuery.Next;
      end;
      columnIndex := columnIndex + 1;
    end;

  end;

end.
