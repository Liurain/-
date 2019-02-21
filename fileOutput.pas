unit fileOutput;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Vcl.OleServer, Word2000, Comobj,DataTransform,StrUtils ;

type
    Output = class
    Constructor Create;
  private
    ado : TAdoOperate;        //���ݿ��������
    ADOQuery : TADOQuery;
    softPath : string;
    visual : boolean; //�Ƿ�Ԥ��



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
    softPath := ExtractFileDir(ParamStr(0));      //��ȡ�����������·���������ļ�����
    visual := false;

    try
      wordApp.Quit;           //�ر���һ��Ԥ��δ�رյĴ��岢�ͷ��ļ�����
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
    modelFile : string;     //ģ���ļ�
    fileName : string;      //�����ļ���



    sql : string;
    i : integer;
  begin
    modelFile := softPath + '\model\��ĩ����.docx';
    fileName:=softPath + '\report\��ĩ����\��ĩ����.docx';
    createWord(modelFile, fileName, wordApp, wordDoc, table, 16);     //��ĩ����ģ��16��

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
                  'A.teacher        as �ον�ʦ' +
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



      //��������
      table.cell(i,1).Range.InsertAfter(ADOQuery.FieldByName('���').AsString);
      table.cell(i,2).Range.InsertAfter(ADOQuery.FieldByName('ʵ��_a').AsString);
      table.cell(i,3).Range.InsertAfter(ADOQuery.FieldByName('ʵ��_b').AsString);
      table.cell(i,4).Range.InsertAfter(ADOQuery.FieldByName('����_a').AsString);
      table.cell(i,5).Range.InsertAfter(ADOQuery.FieldByName('����_b').AsString);
      table.cell(i,6).Range.InsertAfter(ADOQuery.FieldByName('��׼��_a').AsString);
      table.cell(i,7).Range.InsertAfter(ADOQuery.FieldByName('��׼��_b').AsString);
      table.cell(i,8).Range.InsertAfter(ADOQuery.FieldByName('����_a').AsString);
      table.cell(i,9).Range.InsertAfter(ADOQuery.FieldByName('����_b').AsString);
      table.cell(i,10).Range.InsertAfter(ADOQuery.FieldByName('������_a').AsString);
      table.cell(i,11).Range.InsertAfter(ADOQuery.FieldByName('������_b').AsString);
      table.cell(i,12).Range.InsertAfter(ADOQuery.FieldByName('ѧ����_a').AsString);
      table.cell(i,13).Range.InsertAfter(ADOQuery.FieldByName('ѧ����_b').AsString);
      table.cell(i,14).Range.InsertAfter(ADOQuery.FieldByName('�춯_a').AsString);
      table.cell(i,15).Range.InsertAfter(ADOQuery.FieldByName('�춯_b').AsString);
      table.cell(i,16).Range.InsertAfter(ADOQuery.FieldByName('�ον�ʦ').AsString);
      //��ȡ��һ����¼
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
    modelFile : string;     //ģ���ļ�
    fileName : string;      //�����ļ���
    i : integer;
    j : integer;
    dataRows : integer;

    sql : string;
  begin
    modelFile := softPath + '\model\��������.docx';
    fileName:= softPath + '\report\��������\��������_'+examName+'.docx';
    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //��������ģ��9��

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
    if not visual then wordApp.Quit;
  end;

  procedure Output.roster(classID : string;course,exam : string);
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

    trans : CTransform;
    Present: TDateTime;
    Year, Month, Day :Word;
    titleInfo :string;
  begin
    modelFile := softPath + '\model\ѧ���ɼ�ͳ�Ʊ�.docx';
    fileName:= softPath + '\report\�ɼ�ͳ�Ʊ�\ѧ���ɼ�ͳ�Ʊ�'+classID+'.docx';
    if Comparestr('',course) <> 0 then
    begin
      fileName := softPath + '\report\'+exam+'_'+course+'_'+classID+'.docx';
    end;

    createWord(modelFile, fileName, wordApp, wordDoc, table, 9);     //ѧ���ɼ�ͳ�Ʊ�ģ��9��

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
    if not visual then wordApp.Quit;

  end;

  procedure Output.createWord(modelFile, fileName : string;
      var wordApp, wordDoc, table : OleVariant; Columns : integer);
  begin
    CopyFile(PChar(modelFile), PChar(filename), False);      //false��ʾ�滻ͬ�����ļ�

//    try
//      wordApp := GetActiveOleObject('Word.Application');
//    except
      try
        wordApp := CreateOleObject('Word.Application');
        wordApp.Visible := Visual;
      except
        ShowMessage('����word����ʧ�ܣ���ȷ����ʹ�õĵ����Ƿ�װ��word');
        Exit;
      end;
//    end;
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
//    ExcelApp: OleVariant;
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

    createExcel(modelFile, fileName, ExcelApp);

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
      ExcelApp.WorkSheets[1].Cells[i,1].Value := ADOQuery.FieldByName('ѧ��').AsString;
      ExcelApp.WorkSheets[1].Cells[i,2].Value := ADOQuery.FieldByName('����').AsString;
      if Comparestr('',course) <> 0 then
      begin
        ExcelApp.WorkSheets[1].Cells[i,3].Value := ADOQuery.FieldByName('�ɼ�').AsString;
      end;


      //��ȡ��һ����¼
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
    modelFile := softPath + '\model\�ɼ��ܱ�.xlsx';
    fileName := softPath + '\report\�ɼ��ܱ�_'+exam+'_'+classID+'.xlsx';
    createExcel(modelFile, fileName, ExcelApp);

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
      ExcelApp.WorkSheets[1].Cells[i,1].Value := ADOQuery.FieldByName('stuID').AsString;
      ExcelApp.WorkSheets[1].Cells[i,2].Value := ADOQuery.FieldByName('stuName').AsString;
      i := i+1;
      ADOQuery.Next;
    end;

    //�������
    addScoresExcel(classID,exam,courseList,ExcelApp);

    ExcelApp.ActiveWorkBook.Save;
    if not visual then ExcelApp.Quit;
  end;

  procedure Output.createExcel(modelFile, fileName : string;var ExcelApp : OleVariant);
  begin
    CopyFile(PChar(modelFile), PChar(filename), False);      //false��ʾ�滻ͬ�����ļ�

//    try
//      ExcelApp := GetActiveOleObject('Excel.Application');
//    except
      try
        ExcelApp := CreateOleObject('Excel.Application');
        ExcelApp.Visible := Visual;
      except
        ShowMessage('����excel����ʧ�ܣ���ȷ����ʹ�õĵ����Ƿ�װ��excel');
        Exit;
      end;
//    end;
    ExcelApp.WorkBooks.Open(fileName);

    ExcelApp.ActiveSheet.Rows[1].Font.Name := '����';
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
      ExcelApp.WorkSheets[1].Cells[1,columnIndex].Value := courseList[i];  //���ͷ

      sql := 'select stuID, '+exam+' as �ɼ� from tb_scores_'+inttostr(grade)+'_'+courseList[i]+
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
             ExcelApp.WorkSheets[1].Cells[j,columnIndex].Value := ADOQuery.FieldByName('�ɼ�').Asfloat;
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
