unit DatabaseManage;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Table_classes, Table_course, Table_class,
  Table_parameters, Table_users,Table_exam,DataTransform;

type
    Database = class
    Constructor Create;
  private
    ado1 : TAdoOperate;        //���ݿ��������1
    ADOQuery1 : TADOQuery;
    ado2 : TAdoOperate;        //���ݿ��������2
    ADOQuery2 : TADOQuery;
    connString : string;  //��ʱ���ԭ�����ݿ����Ӵ�
    procedure saveConnString();
    procedure setConnString();

  public
    function createNewDatabase(year:integer):boolean;
    procedure keepCourse(year:integer);
    procedure keepPar(year:integer);
    procedure keepTeacher(year:integer);
    procedure keepUser(year:integer);
    procedure keepStu(year:integer);
    procedure initDao(year:integer);
  end;

implementation
  Constructor Database.create;
  begin
    ado1 := TAdoOperate.Create;
    ADOQuery1 := TADOQuery.Create(nil);
    ado2 := TAdoOperate.Create;        //���ݿ��������
    ADOQuery2 := TADOQuery.Create(nil);
  end;

  function Database.createNewDatabase(year:integer):boolean;
  var
    databaseModel : string;
    databaseName : string;
    softPath : string;
    path : string;
    mes : string;  //�Ի�����ʾ��Ϣ
    flag : boolean;
    sql : string;
  begin
    flag := true;
    result := false;
    //�ж��Ƿ����ͬ�����ݿ�
    softPath := ExtractFileDir(ParamStr(0));      //��ȡ�����������·���������ļ�����
    path := softPath + '\database\' + inttostr(year)+'-'+inttostr(year+1);
    if (FileExists(path+'.accdb')) or (FileExists(path+'.mdb')) then
    begin
      mes := inttostr(year)+'-'+inttostr(year+1)+'ѧ�����ݿ��Ѵ��ڣ��Ƿ�ɾ��ԭ�����ݿⴴ�������ݿ⣿';
      if MessageBox(0, PWideChar(mes), '����ѡ��', MB_OKCANCEL + MB_ICONQUESTION) = ID_OK then
      begin
        flag := true;   //�ǣ����ǣ�
      end else begin
        flag := false;
      end;
    end;

    if flag then
    begin
      databaseModel := softPath + '\model\database.accdb';
      databaseName := softPath +'\database\'+inttostr(year)+'-'+inttostr(year+1)+'.accdb';
      CopyFile(PChar(databaseModel), PChar(databaseName), false);      //false��ʾ�滻ͬ�����ļ�
      result := true;

      path := 'database\' + inttostr(year)+'-'+inttostr(year+1)+'.accdb';
      ado1.initConnStr(path);
      if ado1.testConn then
      begin
        sql := 'INSERT INTO tb_parameters(grade,yearBegin) Values(0,' +inttostr(year)+')';
        ado1.ExecSqlStr(ADOQuery1, sql);
        sql := 'INSERT INTO tb_users(users,psw,userIdentity,power) '+
            'Values('+#39+'admin'    +#39+',' +
                      #39+'123'      +#39+',' +
                          '0'            +',' +
                      #39+'111111111'+#39+')';
        ado1.ExecSqlStr(ADOQuery1, sql);
      end;
    end;
  end;

  procedure Database.keepCourse(year:integer);
  var
    sql : string;
    course : Tb_course;
    I: Integer;
  begin
    saveConnString;
    initDao(year);
    course := Tb_course.Create;

    for I := 1 to 9 do
    begin
      sql := 'SELECT courseName,courseType from tb_course_'+inttostr(i);
      ado2.SelectInfo(ADOQuery2, sql);
      while not ADOQuery2.eof do
      begin
        course.grade := i;
        course.courseName := ADOQuery2.FieldByName('courseName').AsString;
        course.courseType := ADOQuery2.FieldByName('courseType').AsString;
        course.addCourse(ADOQuery1);
        ADOQuery2.Next;
      end;
    end;
    setConnString;
  end;

  procedure Database.keepPar(year:integer);
  var
    sql : string;
    averWenBase : real;         //�Ŀ�ѧ������ΪB���𣬵����꼶ƽ���ֵ�������
    averWenTop :real;           //�Ŀ�ѧ������ΪB���𣬸����꼶ƽ���ֵ�������
    averLiBase : real;          //���ѧ������ΪB���𣬵����꼶ƽ���ֵ�������
    averLiTop : real;           //���ѧ������ΪB���𣬸����꼶ƽ���ֵ�������
    dev : real;                 //��׼����ڸò�����ΪA��С����ίC
    greatRatio : real;          //�����ʸ��ڸðٷֱ���ΪA��С�ڸðٷֱ���ίC
    I: Integer;
  begin
    saveConnString;
    initDao(year);

    sql := 'SELECT averWenBase,averWenTop,averLiBase,averLiTop,dev,greatRatio from tb_parameters';
    ado2.SelectInfo(ADOQuery2, sql);

    averWenBase := ADOQuery2.FieldByName('averWenBase').AsFloat;
    averWenTop := ADOQuery2.FieldByName('averWenTop').AsFloat;
    averLiBase := ADOQuery2.FieldByName('averLiBase').AsFloat;
    averLiTop := ADOQuery2.FieldByName('averLiTop').AsFloat;
    dev := ADOQuery2.FieldByName('dev').AsFloat;
    greatRatio := ADOQuery2.FieldByName('greatRatio').AsFloat;

    sql := 'update tb_parameters set '+
        ' averWenBase = ' + floattostr(averWenBase)    + ',' +
        ' averWenTop  = ' + floattostr(averWenTop)     + ',' +
        ' averLiBase  = ' + floattostr(averLiBase)     + ',' +
        ' averLiTop   = ' + floattostr(averLiTop)      + ',' +
        ' dev         = ' + floattostr(dev)            + ',' +
        ' greatRatio  = ' + floattostr(greatRatio)     +
        ' where ID = '+#39+'1'+#39;
    ado1.ExecSqlStr(ADOQuery1, sql);
    setConnString;
  end;

  procedure Database.keepTeacher(year:integer);
  var
    sql : string;
    classes : Tb_classes;
    I: Integer;
  begin
    saveConnString;
    initDao(year);
    classes := Tb_classes.Create;

    sql := 'SELECT classID,stuNum,teacher from tb_classes';
    ado2.SelectInfo(ADOQuery2, sql);
    while not ADOQuery2.eof do
    begin
      try
        classes.classID := ADOQuery2.FieldByName('classID').AsString;
        classes.stuNum := ADOQuery2.FieldByName('stuNum').AsInteger;
        classes.teacher := ADOQuery2.FieldByName('teacher').AsString;
        classes.addClass(ADOQuery1);
      except
        classes.changeClass(ADOQuery1);
      end;
      ADOQuery2.Next;
    end;
    setConnString;
  end;

  procedure Database.keepUser(year:integer);
  var
    sql : string;
    users : string;
    psw : string;
    userIdentity : integer;
    power :string;
    I: Integer;
  begin
    saveConnString;
    initDao(year);

    sql := 'SELECT users,psw,userIdentity,power from tb_users';
    ado2.SelectInfo(ADOQuery2, sql);
    //�������ԭ������
    sql := 'DELETE FROM tb_users WHERE users ='+#39+'admin'+#39;
    ado1.ExecSqlStr(ADOQuery1, sql);
    while not ADOQuery2.eof do
    begin
      users := ADOQuery2.FieldByName('users').AsString;
      psw := ADOQuery2.FieldByName('psw').AsString;
      userIdentity := ADOQuery2.FieldByName('userIdentity').AsInteger;
      power := ADOQuery2.FieldByName('power').AsString;
      sql := 'INSERT INTO tb_users(users,psw,userIdentity,power) Values(' +
          #39 + users             + #39 + ',' +
          #39 + psw               + #39 + ',' +
                inttostr(userIdentity)  + ',' +
          #39 + power             + #39 +')';
      ado1.ExecSqlStr(ADOQuery1, sql);
      ADOQuery2.Next;
    end;
    setConnString;
  end;

  procedure Database.keepStu(year:integer);
  var
    sql : string;
    stu : Tb_class;
    transform : CTransform;
    I : Integer;
    j : integer;
    classIDList : TstringList;
    classes : Tb_classes;
    y : integer;
  begin
    saveConnString;
    initDao(year);

    stu := Tb_class.Create;
    transform := CTransform.Create;
    classIDList := TstringList.Create;
    classes := Tb_classes.Create;

    for I := 1 to 5 do
    begin
      y := transform.GradeToClassID(i);
      sql := 'SELECT classID from tb_classes where classID like '+ inttostr(y) +'& "%"';
      ado2.SelectInfo(ADOQuery2, sql);
      classIDList.clear;
      while not ADOQuery2.eof do
      begin
        classIDList.Add(ADOQuery2.FieldByName('classID').AsString);
        ADOQuery2.Next;
      end;

      for j := 0 to classIDList.Count -1 do
      begin
        try
          classes.classID := classIDList[j];
          classes.stuNum := 0;
          classes.teacher := '';
          classes.addClass(ADOQuery1);
        except
        end;


        sql := 'SELECT * from tb_class_'+classIDList[j];
        ado2.SelectInfo(ADOQuery2, sql);
        while not ADOQuery2.eof do
        begin
          stu.classID := classIDList[j];
          stu.stuID := ADOQuery2.FieldByName('stuID').AsString;
          stu.stuName := ADOQuery2.FieldByName('stuName').AsString;
          stu.sex := ADOQuery2.FieldByName('stuSex').AsString;
          stu.age := ADOQuery2.FieldByName('stuAge').AsInteger;
          stu.addStu(ADOQuery1);
          ADOQuery2.Next;
        end;
      end;
    end;

    setConnString;
  end;

  procedure Database.initDao(year:integer);
  var
    path : string;
    softPath : string;
    newDatabaseName : string;
    oldDatabaseName : string;
  begin
    softPath := ExtractFileDir(ParamStr(0));      //��ȡ�����������·���������ļ�����
    newDatabaseName := inttostr(year)+'-'+inttostr(year+1);
    oldDatabaseName := inttostr(year-1)+'-'+inttostr(year);

    path := softPath + '\database\' + newDatabaseName;
    if (FileExists(path+'.accdb')) then
    begin
      path := 'database\' + newDatabaseName+'.accdb';
    end else if (FileExists(path+'.mdb')) then
    begin
      path := 'database\' + newDatabaseName+'.mdb';
    end;
    ado1.initConnStr(path);       //ȫ�ֱ������Ӵ�

    path := softPath + '\database\' + oldDatabaseName;
    if (FileExists(path+'.accdb')) then
    begin
      path := 'database\' + oldDatabaseName+'.accdb';
    end else if (FileExists(path+'.mdb')) then
    begin
      path := 'database\' + oldDatabaseName+'.mdb';
    end;
    ado2.setConnStr(path);        //�ֲ��������Ӵ�
  end;

  procedure Database.saveConnString();
  begin
    connString := ADOOperate.connStr;
  end;

  procedure Database.setConnString();
  begin
    ADOOperate.connStr := connString;
  end;

end.
