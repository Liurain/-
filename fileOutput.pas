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
    ado : TAdoOperate;        //���ݿ��������
    ADOQuery : TADOQuery;
    softPath : string;
    visual : boolean; //�Ƿ�Ԥ��
    filePrint : boolean; //�Ƿ��ӡ
    wordAppCanClose : boolean;       //wordApp�ܷ�رգ������ж��Ƿ�׷������
    excelAppCanClose : boolean;      //excelApp�ܷ�رգ������ж��Ƿ�׷������



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

{-------------------------------�ɼ�ͳ�Ʊ���----------------------------------}
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
{---------------------------------�ɼ�����------------------------------------}
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
{-------------------------------��ĩ��������----------------------------------}
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
{-------------------------------������������----------------------------------}
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
    softPath := ExtractFileDir(ParamStr(0));      //��ȡ�����������·���������ļ�����
    visual := false;

    try
//      while True do
//      begin
//        wordApp := GetActiveOleObject('Word.Application');
        wordApp.Quit;           //�ر���һ��Ԥ��δ�رյĴ��岢�ͷ��ļ�����
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


{-------------------------------�ɼ�ͳ�Ʊ���----------------------------------}
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
//    modelFile : string;     //ģ���ļ�
//    path : string;          //�ļ����·��
//    fileName : string;      //�����ļ���
//    j : integer;
//    dataRows : integer;
//
//    Present: TDateTime;
//    Year, Month, Day :Word;
//    titleInfo :string;
//  begin
//  //����ģ�沢����word�����ģ��
//    modelFile := softPath + '\model\ȫ�ƵǷֱ�.docx';
//    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
//        'ѧ��\�Ƿֱ�\ȫ�ƵǷֱ�\Word�ļ�';
//    fileName:= path + '\' + classID + 'ȫ�ƵǷֱ�.docx';
//    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //ѧ���ɼ�ͳ�Ʊ�ģ��9��
//  //��ȡ����꼶���пγ�
//    course := Tb_course.Create;
//    trans := CTransform.Create;
//    grade := trans.classIDToGrade(classID);
//    course.getCourseByGrade(ADOQuery, grade);
//    i := 3;   //�ӵ����п�ʼ
//    while not ADOQuery.eof do
//    begin
//      ExcelApp.ActiveSheet.Cells[1,i] := ADOQuery.FieldByName('�γ�����').AsString;
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
  //��ȡ����꼶���еİ༶
    classes.selectByGrade(ADOQuery1, grade);
    while not ADOQuery1.Eof do
    begin
      stat_word_aCourse_class(ADOQuery1.FieldByName('���').AsString);
      ADOQuery1.Next;
    end;
  end;

  procedure Output.stat_word_aCourse_class(classID : string);
  var
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //ģ���ļ�
    path : string;          //�ļ����·��
    fileName : string;      //�����ļ���
    i : integer;
    j : integer;
    dataRows : integer;
    sql : string;

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
  //����ģ�沢����word�����ģ��
    modelFile := softPath + '\model\���ƵǷֱ�.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�Ƿֱ�\���ƵǷֱ�\Word�ļ�';
    fileName:= path + '\' + classID + '���ƵǷֱ�.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //ѧ���ɼ�ͳ�Ʊ�ģ��9��

  //�����ͷ��Ϣ
    trans := CTransform.Create;
    Present:= Now;
    DecodeDate(Present, Year, Month, Day);
    titleInfo := '�༶:'+trans.classIDToText(classID)+'��'+'                     ѧ��:';
    for I := 1 to 26 do
    begin
      titleInfo := titleInfo+' ';
    end;
    titleInfo := titleInfo+inttostr(Year)+'��'+inttostr(Month)+'��'+inttostr(Day)+'��';
    wordDoc.Paragraphs.Item(1).Range.InsertAfter(titleInfo);
  //�����ݿ��л�ȡ����
    sql := 'select stuID          as ѧ��,' +
                  'stuName        as ����' +
              ' from tb_class_'+classID;
    ado.SelectInfo(ADOQuery, sql);
  //��������������������Ӷ�Ӧ������
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
  //�����в�������
    i := 2;     //��ͷռһ�У��ӵڶ��п�ʼ��������
    j := 1;
    while not ADOQuery.eof do
    begin
      //��������
      table.cell(i,j).Range.InsertAfter(ADOQuery.FieldByName('ѧ��').AsString);
      table.cell(i,j+1).Range.InsertAfter(ADOQuery.FieldByName('����').AsString);
      //��ȡ��һ����¼
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�Ƿֱ�\ȫ�ƵǷֱ�\Excel�ļ�';
    fileName:= path + '\' +'�����꼶ѧ���ɼ��Ƿֱ�.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
  //��ȡ���꼶���пγ�
    courseList := TStringList.Create;
    course := Tb_course.Create;
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //�����ظ�ֵ�����
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('�γ�����').AsString);
        ADOQuery.Next;
      end;
    end;

    for I := 0 to courseList.Count -1 do
    begin
      ExcelApp.ActiveSheet.Cells[1,i+3] := courseList[i];
    end;

    //��ѧ�����ݲ���Excel���У��ӵڶ��п�ʼ
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�Ƿֱ�\ȫ�ƵǷֱ�\Excel�ļ�';
    fileName:= path + '\' + inttostr(grade) + '�꼶ѧ���ɼ��Ƿֱ�.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
  //��ȡ����꼶���пγ�
    course := Tb_course.Create;

    course.getCourseByGrade(ADOQuery, grade);
    i := 3;   //�ӵ����п�ʼ
    while not ADOQuery.eof do
    begin
      ExcelApp.ActiveSheet.Cells[1,i] := ADOQuery.FieldByName('�γ�����').AsString;
      i := i+1;
      ADOQuery.Next;
    end;
    //��ѧ�����ݲ���Excel���У��ӵڶ��п�ʼ
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�Ƿֱ�\ȫ�ƵǷֱ�\Excel�ļ�';
    fileName:= path + '\' + classID + '��ѧ���ɼ��Ƿֱ�.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
  //��ȡ����꼶���пγ�
    course := Tb_course.Create;
    trans := CTransform.Create;
    grade := trans.classIDToGrade(classID);

    course.getCourseByGrade(ADOQuery, grade);
    i := 3;   //�ӵ����п�ʼ
    while not ADOQuery.eof do
    begin
      ExcelApp.ActiveSheet.Cells[1,i] := ADOQuery.FieldByName('�γ�����').AsString;
      i := i+1;
      ADOQuery.Next;
    end;
    //��ѧ�����ݲ���Excel���У��ӵڶ��п�ʼ
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�Ƿֱ�\���ƵǷֱ�\Excel�ļ�';
    fileName:= path + '\' +'�����꼶ѧ���ɼ��Ƿֱ�.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //��ѧ�����ݲ���Excel���У��ӵڶ��п�ʼ
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�Ƿֱ�\���ƵǷֱ�\Excel�ļ�';
    fileName:= path + '\' + inttostr(grade) + '�꼶ѧ���ɼ��Ƿֱ�.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //��ѧ�����ݲ���Excel���У��ӵڶ��п�ʼ
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�Ƿֱ�\���ƵǷֱ�\Excel�ļ�';
    fileName:= path + '\' + classID + '��ѧ���ɼ��Ƿֱ�.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //��ѧ�����ݲ���Excel���У��ӵڶ��п�ʼ
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

{---------------------------------�ɼ�����------------------------------------}
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
    t_course : Tb_course;    //course�����
  begin
    t_course := Tb_course.Create;
    for I := 1 to 9 do
    begin
      if t_course.isHasCourse(i, course) then     //�ж�i�꼶�Ƿ�ӵ�пγ�Course
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
  //��ȡ����꼶���еİ༶
    classes.selectByGrade(ADOQuery1, grade);
    while not ADOQuery1.Eof do
    begin
      score_word_aCourse_class(ADOQuery1.FieldByName('���').AsString, course, exam);
      ADOQuery1.Next;
    end;
  end;

  procedure Output.score_word_aCourse_class(classID : string; course : string; exam : string);
  var
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //ģ���ļ�
    path : string;          //�ļ����·��
    fileName : string;      //�����ļ���
    i : integer;
    j : integer;
    dataRows : integer;
    sql : string;

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
    //����ģ�沢����word�����ģ��
    modelFile := softPath + '\model\���ƵǷֱ�.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�ɼ���\���Ƴɼ���\Word�ļ�';
    fileName:= path + '\' + classID + '��ѧ���ɼ���'+exam+'_'+course+'.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //ѧ���ɼ�ͳ�Ʊ�ģ��9��

    //�����ͷ��Ϣ
    trans := CTransform.Create;
    Present:= Now;
    DecodeDate(Present, Year, Month, Day);
    titleInfo := '�༶:'+trans.classIDToText(classID)+'��'+'                     ѧ��:'+course;
    for I := 1 to (26-Length(course)*2) do
    begin
      titleInfo := titleInfo+' ';
    end;
    titleInfo := titleInfo+inttostr(Year)+'��'+inttostr(Month)+'��'+inttostr(Day)+'��';
    wordDoc.Paragraphs.Item(1).Range.InsertAfter(titleInfo);


    //�����ݿ��л�ȡ����
    sql := 'select  A.stuID          as ѧ��,' +
                    ' A.stuName        as ����,' +
                    ' B.'+exam+'       as �ɼ�'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
              ' where A.stuID = B.stuID';
    ado.SelectInfo(ADOQuery, sql);
    //��������������������Ӷ�Ӧ������
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

    i := 2;     //��ͷռһ�У��ӵڶ��п�ʼ��������
    j := 1;
    while not ADOQuery.eof do
    begin
      //��������
      table.cell(i,j).Range.InsertAfter(ADOQuery.FieldByName('ѧ��').AsString);
      table.cell(i,j+1).Range.InsertAfter(ADOQuery.FieldByName('����').AsString);
      table.cell(i,j+2).Range.InsertAfter(ADOQuery.FieldByName('�ɼ�').AsString);
      //��ȡ��һ����¼
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
    //�������ļ������ļ�
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�ɼ���\ȫ�Ƴɼ���\Excel�ļ�';
    fileName:= path + '\' + inttostr(grade) + '�꼶ѧ���ɼ��ܱ�'+exam+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
    //��ȡ���꼶���а༶��Ϣ
    classIDList := TStringList.Create;
    classes := Tb_classes.Create;
    classes.selectByGrade(ADOQuery, grade);
    while not ADOQuery.eof do
    begin
      classIDList.Add(ADOQuery.FieldByName('���').AsString);
      ADOQuery.Next;
    end;

    //��ȡ���꼶���пγ�
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

      //��ʼ��ǰ����
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
      //�������
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�ɼ���\ȫ�Ƴɼ���\Excel�ļ�';
    fileName:= path + '\' + classID + '��ѧ���ɼ��ܱ�'+exam+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //��ȡ���꼶���пγ�
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

    //��ʼ��ǰ����
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


    //�������
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
    t_course : Tb_course;    //course�����
  begin
    t_course := Tb_course.Create;
    for I := 1 to 9 do
    begin
      if t_course.isHasCourse(i, course) then     //�ж�i�꼶�Ƿ�ӵ�пγ�Course
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
  //�������ļ������ļ�
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�ɼ���\���Ƴɼ���\Excel�ļ�';
    fileName:= path + '\' + inttostr(grade) + '�꼶ѧ���ɼ���'+exam+'_'+course+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);
    //��ȡ���꼶���а༶��Ϣ
    classIDList := TStringList.Create;
    classes := Tb_classes.Create;
    classes.selectByGrade(ADOQuery, grade);
    while not ADOQuery.eof do
    begin
      classIDList.Add(ADOQuery.FieldByName('���').AsString);
      ADOQuery.Next;
    end;
    //��Excel�ļ��в���ѧ����Ϣ
    index := 2;      //��ͷռһ�У��ӵڶ��п�ʼ��������
    for I := 0 to classIDList.Count -1 do
    begin
      classID := classIDList[i];
      //�����ݿ��л�ȡ����
      sql := 'select  A.stuID          as ѧ��,' +
                    ' A.stuName        as ����,' +
                    ' B.'+exam+'       as �ɼ�'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(grade)+'_'+course+' B'+
              ' where A.stuID = B.stuID';
      ado.SelectInfo(ADOQuery, sql);

      while not ADOQuery.eof do
      begin
        //��������
        ExcelApp.ActiveSheet.Cells[index,1] := ADOQuery.FieldByName('ѧ��').AsString;
        ExcelApp.ActiveSheet.Cells[index,2] := ADOQuery.FieldByName('����').AsString;
        ExcelApp.ActiveSheet.Cells[index,3] := ADOQuery.FieldByName('�ɼ�').AsString;
        //��ȡ��һ����¼
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
    modelFile := softPath + '\model\�Ƿֱ�.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�ɼ���\���Ƴɼ���\Excel�ļ�';
    fileName:= path + '\' + classID + '��ѧ���ɼ���'+exam+'_'+course+'.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

    //�����ݿ��л�ȡ����
    trans := CTransform.Create;
    sql := 'select  A.stuID          as ѧ��,' +
                  ' A.stuName        as ����,' +
                  ' B.'+exam+'       as �ɼ�'  +
            ' from tb_class_'+classID+ ' A,'+
            ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
            ' where A.stuID = B.stuID';
    trans.Free;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //��ͷռһ�У��ӵڶ��п�ʼ��������
    while not ADOQuery.eof do
    begin
      //��������
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('ѧ��').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('����').AsString;
      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('�ɼ�').AsString;
      //��ȡ��һ����¼
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

{-------------------------------��ĩ��������----------------------------------}
  procedure Output.finalExam_word_allCourse_school(exam : string);
  var
    i : integer;
    courseList : TStringList;
    course : Tb_course;
  begin
    courseList := TStringList.Create;
    course := Tb_course.Create;
    //��ȡ��ѧУ���пγ�
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //�����ظ�ֵ�����
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('�γ�����').AsString);
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
    modelFile : string;     //ģ���ļ�
    path : string;          //
    fileName : string;      //�����ļ���


    sql : string;
    i : integer;

  begin
    //����ģ�沢����word�����ģ��
    modelFile := softPath + '\model\��ĩ����.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\��ĩ����\'+exam+'\Word�ļ�';
    fileName:= path + '\' + course + '��ĩ������.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 16);     //ѧ���ɼ�ͳ�Ʊ�ģ��9��

    //�޸ı���
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.TypeText(course);

//    if comparestr('��ѧ����ĩ',exam) = 0 then
//    begin
//      exam := '��ĩ_��';
//      firstTermFianlExamDataToWord(exam, course, table);
//    end;
//
//    if comparestr('��ѧ����ĩ',exam) = 0 then
//    begin
//      exam := '��ĩ_��';
//      secondTermFianlExamDataToWord(exam, course, table);
//    end;
    exam := '��ĩ';
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
    //��ȡ��ѧУ���пγ�
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //�����ظ�ֵ�����
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('�γ�����').AsString);
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
    modelFile : string;     //ģ���ļ�
    path : string;          //
    fileName : string;      //�����ļ���


    sql : string;
    i : integer;

  begin
    //����ģ�沢����word�����ģ��
    modelFile := softPath + '\model\��ĩ����.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\��ĩ����\'+exam+'\Excel�ļ�';
    fileName:= path + '\' + course + '��ĩ������.xlsx';
    createExcel(modelFile, path, fileName, ExcelApp);

//    if comparestr('��ѧ����ĩ',exam) = 0 then
//    begin
//      exam := '��ĩ_��';
//      FianlExamDataToExcel(exam, course, ExcelApp);
//    end;
//
//    if comparestr('��ѧ����ĩ',exam) = 0 then
//    begin
//      exam := '��ĩ_��';
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
    rowCount : integer;   //�������

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
  //��ȡ��ǰѧ�������������
    ass := Tb_assessment.Create;
    ass.createReport('��ĩ', course);
    //�����ݿ��л�ȡ����
    sql := 'select classID        as ���,' +        //��ѧ�ڰ��
                  'testNum        as ʵ��_b,' +
                  'average        as ����_b,' +
                  'standardDev    as ��׼��_b,' +
                  'greatStuNum    as ����_b,' +
                  'greatRatio     as ������_b,' +
                  'inferiorStuNum as ѧ����_b,' +
                  'evaluate       as �춯_b,' +
                  'teacher        as �ον�ʦ,' +    //��ѧ����ʦ
                  'greatLine      as ������_b,' +
                  'inferiorLine   as ѧ����_b' +
            ' from tb_report_��ĩ_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //��ͷռ���У��ӵ����п�ʼ��������
    while not ADOQuery.eof do
    begin
      //�������������ʱ������
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString +#13+
            '(' + ADOQuery.FieldByName('������_b').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString +#13+
            '(' + ADOQuery.FieldByName('ѧ����_b').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString;
      end;

      //��������
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('���').AsString);
      table.cell(i,2).Range.InsertAfter('0');
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('ʵ��_b').AsString);
      table.cell(i,4).Range.InsertAfter('0');
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('����_b').AsString);
      table.cell(i,6).Range.InsertAfter('0');
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('��׼��_b').AsString);
      table.cell(i,8).Range.InsertAfter('0');
      table.cell(i,9).Range.InsertAfter(greatStuNum_b);
      table.cell(i,10).Range.InsertAfter('0');
      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('������_b').AsString);
      table.cell(i,12).Range.InsertAfter('0');
      table.cell(i,13).Range.InsertAfter(inferiorStuNum_b);
      table.cell(i,14).Range.InsertAfter('');
      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('�춯_b').AsString);
      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('�ον�ʦ').AsString);
      //��ȡ��һ����¼
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //�������


  //��ȡ��һѧ�������������
    //�����һѧ�����ݿ��Ƿ����
    softPath := ExtractFileDir(ParamStr(0));      //��ȡ�����������·���������ļ�����
    if comparestr(GlobalData.term,'��ѧ��') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year-1)+'-'+inttostr(year)+'-��ѧ��';
      yearchange := true;
      moveyear:=1;   //����Ҫ�����꼶����ĩ�������Ǻ��Լ���Ƚϣ�������һѧ��һ�꼶�����ݺͱ�ѧ����꼶��Ӧ
    end else if comparestr(GlobalData.term,'��ѧ��') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year)+'-'+inttostr(year+1)+'-��ѧ��';
      yearchange := false;
      moveyear:=0;   //��ѧ�ں���ѧ����ĩ�Աȣ�û�п�ѧ�꣬���Բ����ٱ������ƶ�һ��
    end;


    if FileExists(oldDatabaseName+'.accdb') then
    begin
      oldDatabaseName := oldDatabaseName + '.accdb';
    end else if FileExists(oldDatabaseName+'.mdb') then begin
      oldDatabaseName := oldDatabaseName + '.mdb';
    end else begin
      exit;                    //û���ҵ���һѧ�����ݿ�
    end;
    //�����������ݿ����Ӵ�
    connString := ADOOperate.connStr;
    //������һѧ�����ݿ�
    ado.initConnStr(oldDatabaseName);
    if ado.testConn then
    begin
      if yearchange then
      begin
        //ע���ʱȫ�ֲ���yearû��
        GlobalData.year := GlobalData.year - 1;
      end;

      //�жϵ�ǰѧ�����ݿ����Ƿ��иÿγ�
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
      //��������ݿ��������ſγ�
      if oldDatabaseHasCourse then
      begin
        ass.createReport('��ĩ', course);
        //�����ݿ��л�ȡ����
        trans := CTransform.Create;
        for I := 3 to rowCount do
        begin
          className := table.cell(i,1).Range.Text;
          className := trans.numToText(trans.TextToNum(LeftStr(className, 1)) - moveyear)
              + RightStr(className, className.Length-1);
          className := LeftStr(className, className.Length-2);
          sql := 'select classID        as ���,' +        //��ѧ�ڰ��
                      'testNum        as ʵ��_a,' +
                      'average        as ����_a,' +
                      'standardDev    as ��׼��_a,' +
                      'greatStuNum    as ����_a,' +
                      'greatRatio     as ������_a,' +
                      'inferiorStuNum as ѧ����_a,' +
                      'evaluate       as �춯_a,' +
                      'greatLine      as ������_a,' +
                      'inferiorLine   as ѧ����_a ' +
                ' from tb_report_��ĩ_' + course +
                ' where classID = ' + #39 + className + #39;
          ado.SelectInfo(ADOQuery, sql);
          //����һ����Ϣ
          if ADOQuery.RecordCount <> 0 then
          begin
            if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
            begin
              greatStuNum_a := ADOQuery.FieldByName('����_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('������_a').AsString + ')';
              inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('ѧ����_a').AsString + ')';
            end else begin
              greatStuNum_a := ADOQuery.FieldByName('����_a').AsString;
              inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString;
            end;

            //��������
            table.cell(i,2).Range.text := ADOQuery.FieldByName('ʵ��_a').AsString;
            table.cell(i,4).Range.text := ADOQuery.FieldByName('����_a').AsString;
            table.cell(i,6).Range.text := ADOQuery.FieldByName('��׼��_a').AsString;
            table.cell(i,8).Range.text := greatStuNum_a;
            table.cell(i,10).Range.text := ADOQuery.FieldByName('������_a').AsString;
            table.cell(i,12).Range.text := inferiorStuNum_a;
            table.cell(i,14).Range.text := ADOQuery.FieldByName('�춯_a').AsString;
          end;
        end;
      end;

      if yearchange then
      begin
        GlobalData.year := GlobalData.year + 1;
      end;
    end else begin
      showMessage('��һѧ�����ݿ�����ʧ��');
    end;
    //�ָ�ԭ�����ݿ�����
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
//    ass.createReport('��ĩ_��', course);
//    ass.createReport('��ĩ_��', course);
//    //�����ݿ��л�ȡ����
//    sql := 'select B.classID        as ���,' +        //��ѧ�ڰ��
//                  'A.testNum        as ʵ��_a,' +
//                  'B.testNum        as ʵ��_b,' +
//                  'A.average        as ����_a,' +
//                  'B.average        as ����_b,' +
//                  'A.standardDev    as ��׼��_a,' +
//                  'B.standardDev    as ��׼��_b,' +
//                  'A.greatStuNum    as ����_a,' +
//                  'B.greatStuNum    as ����_b,' +
//                  'A.greatRatio     as ������_a,' +
//                  'B.greatRatio     as ������_b,' +
//                  'A.inferiorStuNum as ѧ����_a,' +
//                  'B.inferiorStuNum as ѧ����_b,' +
//                  'A.evaluate       as �춯_a,' +
//                  'B.evaluate       as �춯_b,' +
//                  'B.teacher        as �ον�ʦ,' +    //��ѧ����ʦ
//                  'A.greatLine      as ������_a,' +
//                  'B.greatLine      as ������_b,' +
//                  'A.inferiorLine   as ѧ����_a,' +
//                  'B.inferiorLine   as ѧ����_b' +
//            ' from tb_report_��ĩ_��_' + course+' A, tb_report_��ĩ_��_' + course + ' B' +
//            ' where A.classID = B.classID';
//    ado.SelectInfo(ADOQuery, sql);
//
//    i := 3;     //��ͷռ���У��ӵ����п�ʼ��������
//    while not ADOQuery.eof do
//    begin
//      //�������������ʱ������
//      if table.rows.Count < i then
//      begin
//        table.Select;
//        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
//      end;
//
//      if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
//      begin
//        greatStuNum_a := ADOQuery.FieldByName('����_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('������_a').AsString + ')';
//        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('������_b').AsString + ')';
//        inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('ѧ����_a').AsString + ')';
//        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('ѧ����_b').AsString + ')';
//      end else begin
//        greatStuNum_a := ADOQuery.FieldByName('����_a').AsString;
//        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString;
//        inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString;
//        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString;
//      end;
//
//      //��������
//      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('���').AsString);
//      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('ʵ��_a').AsString);
//      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('ʵ��_b').AsString);
//      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('����_a').AsString);
//      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('����_b').AsString);
//      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('��׼��_a').AsString);
//      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('��׼��_b').AsString);
//      table.cell(i,8).Range.InsertAfter(greatStuNum_a);
//      table.cell(i,9).Range.InsertAfter(greatStuNum_b);
//      table.cell(i,10).Range.InsertAfter(ADOQuery.FieldByName('������_a').AsString);
//      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('������_b').AsString);
//      table.cell(i,12).Range.InsertAfter(inferiorStuNum_a);
//      table.cell(i,13).Range.InsertAfter(inferiorStuNum_b);
//      table.cell(i,14).Range.InsertAfter(ADOQuery.FieldByName('�춯_a').AsString);
//      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('�춯_b').AsString);
//      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('�ον�ʦ').AsString);
//      //��ȡ��һ����¼
//      ADOQuery.Next;
//      i := i+1;
//    end;
  end;

  procedure Output.FianlExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
  var
    sql : string;
    I : integer;
    rowCount : integer;   //�������

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
  //��ȡ��ǰѧ�������������
    ass := Tb_assessment.Create;
    ass.createReport('��ĩ', course);
    //�����ݿ��л�ȡ����
    sql := 'select classID        as ���,' +        //��ѧ�ڰ��
                  'testNum        as ʵ��_b,' +
                  'average        as ����_b,' +
                  'standardDev    as ��׼��_b,' +
                  'greatStuNum    as ����_b,' +
                  'greatRatio     as ������_b,' +
                  'inferiorStuNum as ѧ����_b,' +
                  'evaluate       as �춯_b,' +
                  'teacher        as �ον�ʦ,' +    //��ѧ����ʦ
                  'greatLine      as ������_b,' +
                  'inferiorLine   as ѧ����_b' +
            ' from tb_report_��ĩ_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //��ͷռ���У��ӵ����п�ʼ��������
    while not ADOQuery.eof do
    begin
      if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString +#13+
            '(' + ADOQuery.FieldByName('������_b').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString +#13+
            '(' + ADOQuery.FieldByName('ѧ����_b').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString;
      end;

      //��������
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('���').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := '0';
      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('ʵ��_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,4] := '0';
      ExcelApp.ActiveSheet.Cells[i,5] := ADOQuery.FieldByName('����_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,6] := '0';
      ExcelApp.ActiveSheet.Cells[i,7] := ADOQuery.FieldByName('��׼��_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,8] := '0';
      ExcelApp.ActiveSheet.Cells[i,9] := greatStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,10] := '0';
      ExcelApp.ActiveSheet.Cells[i,11] := ADOQuery.FieldByName('������_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,12] := '0';
      ExcelApp.ActiveSheet.Cells[i,13] := inferiorStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,14] := '';
      ExcelApp.ActiveSheet.Cells[i,15] := ADOQuery.FieldByName('�춯_b').AsString;
      ExcelApp.ActiveSheet.Cells[i,16] := ADOQuery.FieldByName('�ον�ʦ').AsString;
      //��ȡ��һ����¼
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //�������


  //��ȡ��һѧ�������������
    //�����һѧ�����ݿ��Ƿ����
    softPath := ExtractFileDir(ParamStr(0));      //��ȡ�����������·���������ļ�����
    if comparestr(GlobalData.term,'��ѧ��') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year-1)+'-'+inttostr(year)+'-��ѧ��';
      yearchange := true;
      moveyear:=1;   //����Ҫ�����꼶����ĩ�������Ǻ��Լ���Ƚϣ�������һѧ��һ�꼶�����ݺͱ�ѧ����꼶��Ӧ
    end else if comparestr(GlobalData.term,'��ѧ��') = 0 then
    begin
      oldDatabaseName := softPath + '\database\'+ inttostr(year)+'-'+inttostr(year+1)+'-��ѧ��';
      yearchange := false;
      moveyear:=0;   //��ѧ�ں���ѧ����ĩ�Աȣ�û�п�ѧ�꣬���Բ����ٱ������ƶ�һ��
    end;


    if FileExists(oldDatabaseName+'.accdb') then
    begin
      oldDatabaseName := oldDatabaseName + '.accdb';
    end else if FileExists(oldDatabaseName+'.mdb') then begin
      oldDatabaseName := oldDatabaseName + '.mdb';
    end else begin
      exit;                    //û���ҵ���һѧ�����ݿ�
    end;
    //�����������ݿ����Ӵ�
    connString := ADOOperate.connStr;
    //������һѧ�����ݿ�
    ado.initConnStr(oldDatabaseName);
    if ado.testConn then
    begin
      if yearchange then
      begin
        //ע���ʱȫ�ֲ���yearû��
        GlobalData.year := GlobalData.year - 1;
      end;

      //�жϵ�ǰѧ�����ݿ����Ƿ��иÿγ�
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
      //��������ݿ��������ſγ�
      if oldDatabaseHasCourse then
      begin
        ass.createReport('��ĩ', course);
        //�����ݿ��л�ȡ����
        trans := CTransform.Create;
        for I := 3 to rowCount do
        begin
          className := ExcelApp.ActiveSheet.Cells[i,1];
          className := trans.numToText(trans.TextToNum(LeftStr(className, 1)) - 1)
              + RightStr(className, className.Length-1);
          //className := LeftStr(className, className.Length-2);
          sql := 'select classID        as ���,' +        //��ѧ�ڰ��
                      'testNum        as ʵ��_a,' +
                      'average        as ����_a,' +
                      'standardDev    as ��׼��_a,' +
                      'greatStuNum    as ����_a,' +
                      'greatRatio     as ������_a,' +
                      'inferiorStuNum as ѧ����_a,' +
                      'evaluate       as �춯_a,' +
                      'greatLine      as ������_a,' +
                      'inferiorLine   as ѧ����_a ' +
                ' from tb_report_��ĩ_' + course +
                ' where classID = ' + #39 + className + #39;
          ado.SelectInfo(ADOQuery, sql);
          //����һ����Ϣ
          if ADOQuery.RecordCount <> 0 then
          begin
            if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
            begin
              greatStuNum_a := ADOQuery.FieldByName('����_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('������_a').AsString + ')';
              inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString +#13+
                  '(' + ADOQuery.FieldByName('ѧ����_a').AsString + ')';
            end else begin
              greatStuNum_a := ADOQuery.FieldByName('����_a').AsString;
              inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString;
            end;

            //��������
            ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('ʵ��_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,4] :=  ADOQuery.FieldByName('����_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,6] := ADOQuery.FieldByName('��׼��_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,8] := greatStuNum_a;
            ExcelApp.ActiveSheet.Cells[i,10] := ADOQuery.FieldByName('������_a').AsString;
            ExcelApp.ActiveSheet.Cells[i,12] := inferiorStuNum_a;
            ExcelApp.ActiveSheet.Cells[i,14] := ADOQuery.FieldByName('�춯_a').AsString;
          end;
        end;
      end;

      if yearchange then
      begin
        //ע���ʱȫ�ֲ���yearû��
        GlobalData.year := GlobalData.year + 1;
      end;
    end else begin
      showMessage('��һѧ�����ݿ�����ʧ��');
    end;
    //�ָ�ԭ�����ݿ�����
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
//    ass.createReport('��ĩ_��', course);
//    ass.createReport('��ĩ_��', course);
//    //�����ݿ��л�ȡ����
//    sql := 'select B.classID        as ���,' +        //��ѧ�ڰ��
//                  'A.testNum        as ʵ��_a,' +
//                  'B.testNum        as ʵ��_b,' +
//                  'A.average        as ����_a,' +
//                  'B.average        as ����_b,' +
//                  'A.standardDev    as ��׼��_a,' +
//                  'B.standardDev    as ��׼��_b,' +
//                  'A.greatStuNum    as ����_a,' +
//                  'B.greatStuNum    as ����_b,' +
//                  'A.greatRatio     as ������_a,' +
//                  'B.greatRatio     as ������_b,' +
//                  'A.inferiorStuNum as ѧ����_a,' +
//                  'B.inferiorStuNum as ѧ����_b,' +
//                  'A.evaluate       as �춯_a,' +
//                  'B.evaluate       as �춯_b,' +
//                  'B.teacher        as �ον�ʦ,' +    //��ѧ����ʦ
//                  'A.greatLine      as ������_a,' +
//                  'B.greatLine      as ������_b,' +
//                  'A.inferiorLine   as ѧ����_a,' +
//                  'B.inferiorLine   as ѧ����_b' +
//            ' from tb_report_��ĩ_��_' + course+' A, tb_report_��ĩ_��_' + course + ' B' +
//            ' where A.classID = B.classID';
//    ado.SelectInfo(ADOQuery, sql);
//
//    i := 3;     //��ͷռ���У��ӵ����п�ʼ��������
//    while not ADOQuery.eof do
//    begin
//      if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
//      begin
//        greatStuNum_a := ADOQuery.FieldByName('����_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('������_a').AsString + ')';
//        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('������_b').AsString + ')';
//        inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString +#13+
//            '(' + ADOQuery.FieldByName('ѧ����_a').AsString + ')';
//        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString +#13+
//            '(' + ADOQuery.FieldByName('ѧ����_b').AsString + ')';
//      end else begin
//        greatStuNum_a := ADOQuery.FieldByName('����_a').AsString;
//        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString;
//        inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString;
//        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString;
//      end;
//
//      //��������
//      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('���').AsString;
//      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('ʵ��_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('ʵ��_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,4] := ADOQuery.FieldByName('����_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,5] := ADOQuery.FieldByName('����_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,6] := ADOQuery.FieldByName('��׼��_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,7] := ADOQuery.FieldByName('��׼��_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,8] := greatStuNum_a;
//      ExcelApp.ActiveSheet.Cells[i,9] := greatStuNum_b;
//      ExcelApp.ActiveSheet.Cells[i,10] := ADOQuery.FieldByName('������_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,11] := ADOQuery.FieldByName('������_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,12] := inferiorStuNum_a;
//      ExcelApp.ActiveSheet.Cells[i,13] := inferiorStuNum_b;
//      ExcelApp.ActiveSheet.Cells[i,14] := ADOQuery.FieldByName('�춯_a').AsString;
//      ExcelApp.ActiveSheet.Cells[i,15] := ADOQuery.FieldByName('�춯_b').AsString;
//      ExcelApp.ActiveSheet.Cells[i,16] := ADOQuery.FieldByName('�ον�ʦ').AsString;
//      //��ȡ��һ����¼
//      ADOQuery.Next;
//      i := i+1;
//    end;
  end;

{-------------------------------������������----------------------------------}
  procedure Output.exam_word_allCourse_school( exam : string);
  var
    i : integer;
    courseList : TStringList;
    course : Tb_course;
  begin
    courseList := TStringList.Create;
    course := Tb_course.Create;
    //��ȡ��ѧУ���пγ�
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //�����ظ�ֵ�����
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('�γ�����').AsString);
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
    modelFile : string;     //ģ���ļ�
    path : string;          //
    fileName : string;      //�����ļ���


    sql : string;
    i : integer;

  begin
    //����ģ�沢����word�����ģ��
    modelFile := softPath + '\model\��������.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\��������\'+exam+'\Word�ļ�';
    fileName:= path + '\' +exam+'-'+ course + '-����������.docx';
    createWord(modelFile, path, fileName, wordApp, wordDoc, table, 9);     //ѧ���ɼ�ͳ�Ʊ�ģ��9��

    //�޸ı���
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
    //��ȡ��ѧУ���пγ�
    courseList.Sorted := true;
    courseList.Duplicates := dupIgnore;  //�����ظ�ֵ�����
    for I := 1 to 9 do
    begin
      course.getCourseByGrade(ADOQuery, i);
      while not ADOQuery.eof do
      begin
        courseList.Add(ADOQuery.FieldByName('�γ�����').AsString);
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
    modelFile : string;     //ģ���ļ�
    path : string;          //
    fileName : string;      //�����ļ���


    sql : string;
    i : integer;

  begin
    //����ģ�沢����word�����ģ��
    modelFile := softPath + '\model\��������.xlsx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\��������\'+exam+'\Excel�ļ�';
    fileName:= path + '\'+exam+'-' + course + '-����������.xlsx';
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
    rowCount : integer;   //�������

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
  //��ȡ��ǰѧ�������������
    ass := Tb_assessment.Create;
    ass.createReport(exam, course);
    //�����ݿ��л�ȡ����
    sql := 'select classID        as ���,' +
                  'testNum        as ʵ��,' +
                  'average        as ����,' +
                  'standardDev    as ��׼��,' +
                  'greatStuNum    as ����,' +
                  'greatRatio     as ������,' +
                  'inferiorStuNum as ѧ����,' +
                  'evaluate       as �춯,' +
                  'teacher        as �ον�ʦ,' +
                  'greatLine      as ������,' +
                  'inferiorLine   as ѧ����' +
            ' from tb_report_'+exam+'_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //��ͷռ���У��ӵ����п�ʼ��������
    while not ADOQuery.eof do
    begin
      //�������������ʱ������
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('����').AsString +#13+
            '(' + ADOQuery.FieldByName('������').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����').AsString +#13+
            '(' + ADOQuery.FieldByName('ѧ����').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('����').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����').AsString;
      end;

      //��������
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('���').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('ʵ��').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('����').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('��׼��').AsString);
      table.cell(i,5).Range.InsertAfter(greatStuNum_b);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('������').AsString);
      table.cell(i,7).Range.InsertAfter(inferiorStuNum_b);
      table.cell(i,8).Range.InsertAfter(ADOQuery.FieldByName('�춯').AsString);
      table.cell(i,9).Range.InsertAfter(ADOQuery.FieldByName('�ον�ʦ').AsString);
      //��ȡ��һ����¼
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //�������
  end;


  procedure Output.ExamDataToExcel(exam : string; course : string; var ExcelApp: OleVariant);
  var
    sql : string;
    I : integer;
    rowCount : integer;   //�������

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
  //��ȡ��ǰѧ�������������
    ass := Tb_assessment.Create;
    ass.createReport(exam, course);
    //�����ݿ��л�ȡ����
    sql := 'select classID        as ���,' +
                  'testNum        as ʵ��,' +
                  'average        as ����,' +
                  'standardDev    as ��׼��,' +
                  'greatStuNum    as ����,' +
                  'greatRatio     as ������,' +
                  'inferiorStuNum as ѧ����,' +
                  'evaluate       as �춯,' +
                  'teacher        as �ον�ʦ,' +
                  'greatLine      as ������,' +
                  'inferiorLine   as ѧ����' +
            ' from tb_report_'+exam+'_' + course;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //��ͷռ���У��ӵڶ��п�ʼ��������
    while not ADOQuery.eof do
    begin
      if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
      begin
        greatStuNum_b := ADOQuery.FieldByName('����').AsString +#13+
            '(' + ADOQuery.FieldByName('������').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����').AsString +#13+
            '(' + ADOQuery.FieldByName('ѧ����').AsString + ')';
      end else begin
        greatStuNum_b := ADOQuery.FieldByName('����').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����').AsString;
      end;

      //��������
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('���').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('ʵ��').AsString;
      ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('����').AsString;
      ExcelApp.ActiveSheet.Cells[i,4] := ADOQuery.FieldByName('��׼��').AsString;
      ExcelApp.ActiveSheet.Cells[i,5] := greatStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,6] := ADOQuery.FieldByName('������').AsString;
      ExcelApp.ActiveSheet.Cells[i,7] := inferiorStuNum_b;
      ExcelApp.ActiveSheet.Cells[i,8] := ADOQuery.FieldByName('�춯').AsString;
      ExcelApp.ActiveSheet.Cells[i,9] := ADOQuery.FieldByName('�ον�ʦ').AsString;
      //��ȡ��һ����¼
      ADOQuery.Next;
      i := i+1;
    end;
    rowCount := i;   //�������
  end;

{***************************************Word**************************************}
  procedure Output.finalreport(course : string);
  var
//    wordApp: OleVariant;
    wordDoc: OleVariant;
    table: OleVariant;
    modelFile : string;     //ģ���ļ�
    fileName : string;      //�����ļ���

    greatStuNum_a : string;
    greatStuNum_b : string;
    inferiorStuNum_a : string;
    inferiorStuNum_b : string;

    sql : string;
    i : integer;
  begin
    modelFile := softPath + '\model\��ĩ����.docx';
    fileName:=softPath + '\report\��ĩ����\��ĩ����_'
        +course+inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1)+'.docx';
//    createWord(modelFile, fileName, wordApp, wordDoc, table, 16);     //��ĩ����ģ��16��

    //�޸ı���
    wordApp.Selection.delete;
    wordApp.Selection.delete;
    wordApp.Selection.TypeText(course);

    //�����ݿ��л�ȡ����
    sql := 'select A.classID        as ���,' +
                  'A.testNum        as ʵ��_a,' +
                  'B.testNum        as ʵ��_b,' +
                  'A.average        as ����_a,' +
                  'B.average        as ����_b,' +
                  'A.standardDev    as ��׼��_a,' +
                  'B.standardDev    as ��׼��_b,' +
                  'A.greatStuNum    as ����_a,' +
                  'B.greatStuNum    as ����_b,' +
                  'A.greatRatio     as ������_a,' +
                  'B.greatRatio     as ������_b,' +
                  'A.inferiorStuNum as ѧ����_a,' +
                  'B.inferiorStuNum as ѧ����_b,' +
                  'A.evaluate       as �춯_a,' +
                  'B.evaluate       as �춯_b,' +
                  'A.teacher        as �ον�ʦ,' +
                  'A.greatLine      as ������_a,' +
                  'B.greatLine      as ������_b,' +
                  'A.inferiorLine   as ѧ����_a,' +
                  'B.inferiorLine   as ѧ����_b' +
            ' from tb_report_��ĩ_��_'+course+' A,tb_report_��ĩ_��_'+course+' B' +
            ' where A.classID = B.classID';
    ado.SelectInfo(ADOQuery, sql);

    i := 3;     //��ͷռ���У��ӵ����п�ʼ��������
    while not ADOQuery.eof do
    begin
      //�������������ʱ������
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      if comparestr('�꼶',RightStr(ADOQuery.FieldByName('���').AsString,2)) = 0 then
      begin
        greatStuNum_a := ADOQuery.FieldByName('����_a').AsString +#13+
            '(' + ADOQuery.FieldByName('������_a').AsString + ')';
        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString +#13+
            '(' + ADOQuery.FieldByName('������_b').AsString + ')';
        inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString +#13+
            '(' + ADOQuery.FieldByName('ѧ����_a').AsString + ')';
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString +#13+
            '(' + ADOQuery.FieldByName('ѧ����_b').AsString + ')';
      end else begin
        greatStuNum_a := ADOQuery.FieldByName('����_a').AsString;
        greatStuNum_b := ADOQuery.FieldByName('����_b').AsString;
        inferiorStuNum_a := ADOQuery.FieldByName('ѧ����_a').AsString;
        inferiorStuNum_b := ADOQuery.FieldByName('ѧ����_b').AsString;
      end;


      //��������
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('���').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('ʵ��_a').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('ʵ��_b').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('����_a').AsString);
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('����_b').AsString);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('��׼��_a').AsString);
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('��׼��_b').AsString);
      table.cell(i,8).Range.InsertAfter(greatStuNum_a);
      table.cell(i,9).Range.InsertAfter(greatStuNum_b);
      table.cell(i,10).Range.InsertAfter(ADOQuery.FieldByName('������_a').AsString);
      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('������_b').AsString);
      table.cell(i,12).Range.InsertAfter(inferiorStuNum_a);
      table.cell(i,13).Range.InsertAfter(inferiorStuNum_b);
      table.cell(i,14).Range.InsertAfter(ADOQuery.FieldByName('�춯_a').AsString);
      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('�춯_b').AsString);
      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('�ον�ʦ').AsString);
      //��ȡ��һ����¼
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
    modelFile : string;     //ģ���ļ�
    fileName : string;      //�����ļ���
    i : integer;
    j : integer;
    dataRows : integer;

    sql : string;
  begin
    modelFile := softPath + '\model\��������.docx';
    fileName:= softPath + '\report\��������\��������_'+examName+'.docx';
//    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //��������ģ��9��

    //�޸ı���
    for I := 1 to 6 do
    begin
      wordApp.Selection.delete;
    end;
    wordApp.Selection.TypeText(course+'ѧ��'+examName);

    //�����ݿ��л�ȡ����
    sql := 'select classID        as ���,' +
                  'testNum        as ʵ��,' +
                  'average        as ����,' +
                  'standardDev    as ��׼��,' +
                  'greatStuNum    as ����,' +
                  'greatRatio     as ������,' +
                  'inferiorStuNum as ѧ����,' +
                  'evaluate       as �춯,' +
                  'teacher        as �ον�ʦ' +
            ' from tb_report_'+examName+'_'+course;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //��ͷռһ�У��ӵڶ��п�ʼ��������
    while not ADOQuery.eof do
    begin
      //�������������ʱ������
      if table.rows.Count < i then
      begin
        table.Select;
        wordApp.Selection.InsertRowsBelow(i-table.rows.Count);
      end;

      //��������
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('���').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('ʵ��').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('����').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('��׼��').AsString);
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('����').AsString);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('������').AsString);
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('ѧ����').AsString);
      table.cell(i,8).Range.InsertAfter(ADOQuery.FieldByName('�춯').AsString);
      table.cell(i,9).Range.InsertAfter(ADOQuery.FieldByName('�ον�ʦ').AsString);
      //��ȡ��һ����¼
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
    modelFile : string;     //ģ���ļ�
    path : string;          //�ļ����·��
    fileName : string;      //�����ļ���
    i : integer;
    j : integer;
    dataRows : integer;
    sql : string;

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
    modelFile := softPath + '\model\ѧ���ɼ�ͳ�Ʊ�.docx';
    path := softPath + '\report\' + inttostr(GlobalData.year)+'-'+inttostr(GlobalData.year+1) +
        'ѧ��-'+GlobalData.term+'\�ɼ�ͳ�Ʊ�';
    fileName:= path + '\' + classID + '��ѧ���ɼ�ͳ�Ʊ�.docx';
    if Comparestr('',course) <> 0 then
    begin
      fileName := softPath + '\report\'+exam+'_'+course+'_'+classID+'.docx';
    end;

//    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //ѧ���ɼ�ͳ�Ʊ�ģ��9��

    //�����ͷ��Ϣ
    trans := CTransform.Create;
    Present:= Now;
    DecodeDate(Present, Year, Month, Day);
    titleInfo := '�༶:'+trans.classIDToText(classID)+'��'+'                     ѧ��:'+course;

    for I := 1 to (26-Length(course)*2) do
    begin
      titleInfo := titleInfo+' ';
    end;

    titleInfo := titleInfo+inttostr(Year)+'��'+inttostr(Month)+'��'+inttostr(Day)+'��';

    wordDoc.Paragraphs.Item(1).Range.InsertAfter(titleInfo);


    //�����ݿ��л�ȡ����
    if Comparestr('',course) = 0 then
    begin
      sql := 'select stuID          as ѧ��,' +
                    'stuName        as ����' +
              ' from tb_class_'+classID;
    end else begin
      sql := 'select  A.stuID          as ѧ��,' +
                    ' A.stuName        as ����,' +
                    ' B.'+exam+'       as �ɼ�'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
              ' where A.stuID = B.stuID';
    end;
    ado.SelectInfo(ADOQuery, sql);
    //��������������������Ӷ�Ӧ������
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

    i := 2;     //��ͷռһ�У��ӵڶ��п�ʼ��������
    j := 1;
    while not ADOQuery.eof do
    begin
      //��������
      table.cell(i,j).Range.InsertAfter(ADOQuery.FieldByName('ѧ��').AsString);
      table.cell(i,j+1).Range.InsertAfter(ADOQuery.FieldByName('����').AsString);
      if Comparestr('',course) <> 0 then
      begin
        table.cell(i,j+2).Range.InsertAfter(ADOQuery.FieldByName('�ɼ�').AsString);
      end;
      //��ȡ��һ����¼
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
  //�ж��ļ�·���Ƿ����,�����������ȴ�������ļ���
    if not DirectoryExists(path) then //�ж�Ŀ¼�Ƿ����
    try
      begin
        ForceDirectories(path);   //����Ŀ¼
      end;
    except
      raise Exception.Create('�޷�����·���� '+path);
    end;
  //����ģ���ļ�
    CopyFile(PChar(modelFile), PChar(filename), False);      //false��ʾ�滻ͬ�����ļ�
  //����word���󣬲����ļ�
    try
      wordApp := CreateOleObject('Word.Application');
      wordApp.Visible := false;
    except
      ShowMessage('����word����ʧ�ܣ���ȷ����ʹ�õĵ����Ƿ�װ��word');
      Exit;
    end;
    wordDoc := WordApp.Documents.Open(filename,false,false,false);
    //���ģ���Ƿ���Ϲ淶
    try
      table := WordDoc.Tables.Item(1);
      if table.Columns.Count <> Columns then
      begin
        ShowMessage('ģ���б�񲻷��Ϲ淶��Ӧ����'+inttostr(Columns)+'�У�');
        wordApp.Quit;
        Exit;
      end;
    except
      ShowMessage('ģ���б��ʧ��');
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
    modelFile := softPath + '\model\ѧ���ɼ�ͳ�Ʊ�.xlsx';
    fileName := softPath + '\report\ѧ���ɼ�ͳ�Ʊ�_'+classID+'.xlsx';
    if Comparestr('',course) <> 0 then
    begin
      fileName := softPath + '\report\'+exam+'_'+course+'_'+classID+'.xlsx';
    end;

//    createExcel(modelFile, fileName, ExcelApp);

    //�����ݿ��л�ȡ����
    if Comparestr('',course) = 0 then
    begin
      sql := 'select stuID          as ѧ��,' +
                    'stuName        as ����' +
              ' from tb_class_'+classID;
    end else begin
      trans := CTransform.Create;
      sql := 'select  A.stuID          as ѧ��,' +
                    ' A.stuName        as ����,' +
                    ' B.'+exam+'       as �ɼ�'  +
              ' from tb_class_'+classID+ ' A,'+
              ' tb_scores_'+inttostr(trans.classIDToGrade(classID))+'_'+course+' B'+
              ' where A.stuID = B.stuID';
      trans.Free;
    end;
    ado.SelectInfo(ADOQuery, sql);

    i := 2;     //��ͷռһ�У��ӵڶ��п�ʼ��������
    while not ADOQuery.eof do
    begin
      //��������
      ExcelApp.ActiveSheet.Cells[i,1] := ADOQuery.FieldByName('ѧ��').AsString;
      ExcelApp.ActiveSheet.Cells[i,2] := ADOQuery.FieldByName('����').AsString;
      if Comparestr('',course) <> 0 then
      begin
        ExcelApp.ActiveSheet.Cells[i,3] := ADOQuery.FieldByName('�ɼ�').AsString;
      end;


      //��ȡ��һ����¼
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
    modelFile := softPath + '\model\�ɼ��ܱ�.xlsx';
    fileName := softPath + '\report\�ɼ��ܱ�_'+exam+'_'+classID+'.xlsx';
//    createExcel(modelFile, fileName, ExcelApp);

    //��ȡ���꼶���пγ�
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

    //��ʼ��ǰ����
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

    //�������
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
  //�ж��ļ�·���Ƿ����,�����������ȴ�������ļ���
    if not DirectoryExists(path) then //�ж�Ŀ¼�Ƿ����
    try
      begin
        ForceDirectories(path);   //����Ŀ¼
      end;
    except
      raise Exception.Create('�޷�����·���� '+path);
    end;
    CopyFile(PChar(modelFile), PChar(filename), False);      //false��ʾ�滻ͬ�����ļ�

    try
      ExcelApp := CreateOleObject('Excel.Application');
      ExcelApp.Visible := false;
    except
      ShowMessage('����excel����ʧ�ܣ���ȷ����ʹ�õĵ����Ƿ�װ��excel');
      Exit;
    end;
    ExcelApp.WorkBooks.Open(fileName);

    ExcelApp.ActiveSheet.Rows[1].Font.Name := '����';
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
      ExcelApp.ActiveSheet.Cells[1,columnIndex] := courseList[i];  //���ͷ

      sql := 'select stuID, '+exam+' as �ɼ� from tb_scores_'+inttostr(grade)+'_'+courseList[i]+
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
             ExcelApp.ActiveSheet.Cells[j,columnIndex] := ADOQuery.FieldByName('�ɼ�').Asfloat;
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
    //�����ݿ��л�ȡ����
    sql := 'select stuID      as ѧ��,' +
              ' stuName       as ����' +
              ' from tb_class_'+classID;
    ado.SelectInfo(ADOQuery, sql);

    while not ADOQuery.eof do
    begin
      //��������
      ExcelApp.ActiveSheet.Cells[index,1] := ADOQuery.FieldByName('ѧ��').AsString;
      ExcelApp.ActiveSheet.Cells[index,2] := ADOQuery.FieldByName('����').AsString;
      //��ȡ��һ����¼
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
      index := insterClassStuInfoToExcel(index,ADOQuery1.FieldByName('���').AsString);
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
