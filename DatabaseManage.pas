unit DatabaseManage;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Table_classes, Table_course, Table_class,
  Table_parameters, Table_users,Table_exam, DataTransform;

type
    Database = class
    Constructor Create;
  private
    ado : TAdoOperate;        //数据库操作对象1
    ADOQuery : TADOQuery;
    connString : string;  //临时存放原有数据库连接串

    function copyDatabase(year:integer; term:string; databaseModel : string):boolean;
  public
    function createNewDatabase(year:integer;term:string):boolean;
    function createDatabaseByOld(year:integer; term:string):boolean;


    procedure saveConnString();
    function linkNewDatabase(year:integer; term:string):boolean;
    procedure setBeginYear(year : integer);
    procedure recoverOldLink();

    procedure delCourse();
    procedure delPar();
    procedure delTeacher();
    procedure delUser();
    procedure delStu(grade : integer);
    procedure updateCourseTable();
    procedure delExam(keepStu:boolean);
  end;

implementation
  Constructor Database.create;
  begin
    ado := TAdoOperate.Create;
    ADOQuery := TADOQuery.Create(nil);
  end;

  function Database.createNewDatabase(year:integer; term:string):boolean;
  var
    databaseModel : string;
    softPath : string;
    path : string;

    sql : string;
  begin
    softPath := ExtractFileDir(ParamStr(0));      //获取软件本身所在路径，用于文件操作
    databaseModel := softPath + '\model\database.accdb';
    result := copyDatabase(year,term, databaseModel);

    path := 'database\' + inttostr(year)+'-'+inttostr(year+1)+'-'+term+'.accdb';
    ado.initConnStr(path);
    if ado.testConn then
    begin
      sql := 'update tb_parameters set '+
        ' yearBegin = ' + inttostr(year) +
        ' where grade = 0';
      ado.ExecSqlStr(sql);
      try
        sql := 'INSERT INTO tb_users(users,psw,userIdentity,power) '+
          'Values('+#39+'admin'    +#39+',' +
                    #39+'123'      +#39+',' +
                        '0'            +',' +
                    #39+'111111111'+#39+')';
        ado.ExecSqlStr(sql);
      except

      end;
    end;
    ado.closeConn;
  end;

  function Database.createDatabaseByOld(year:integer; term:string):boolean;
  var
    databaseModel : string;
    softPath : string;
    path : string;

    sql : string;
  begin
    //复制旧的数据库
    softPath := ExtractFileDir(ParamStr(0));      //获取软件本身所在路径，用于文件操作
    if comparestr(term,'上学期') = 0 then
    begin
      databaseModel := softPath + '\database\'+ inttostr(year-1)+'-'+inttostr(year)+'-下学期';
    end else if comparestr(term,'下学期') = 0 then
    begin
      databaseModel := softPath + '\database\'+ inttostr(year)+'-'+inttostr(year+1)+'-上学期';
    end;
    if FileExists(databaseModel+'.accdb') then
    begin
      databaseModel := databaseModel + '.accdb';
    end else if FileExists(databaseModel+'.mdb') then begin
      databaseModel := databaseModel + '.mdb';
    end else begin
      showmessage('所需参照的旧数据库'+databaseModel+'不存在,将创建一个空白的新数据库。');
      databaseModel := softPath + '\model\database.accdb';
//      result := false;
    end;
    result := copyDatabase(year,term, databaseModel);
    //删除旧的分析表
//    sql := 'Drop table tb_report& "%"';
//    ado.ExecSqlStr(sql);
  end;

  procedure Database.recoverOldLink();
  var
    par : Tb_parameters;
  begin
    ado.setConnStr(connString);
    par := Tb_parameters.Create;
    par.getPar();
  end;

  function Database.linkNewDatabase(year:integer; term:string):boolean;
  var
    path : string;
  begin
    path := 'database\' + inttostr(year)+'-'+inttostr(year+1)+'-'+term;
    if FileExists(path+'.accdb') then
    begin
      path := path + '.accdb';
    end else if FileExists(path+'.mdb') then begin
      path := path + '.mdb';
    end;

    ado.initConnStr(path);
    if ado.testConn then
    begin
      result := true;
    end else begin
      result := false;
    end;

    //跨学年需要清除毕业班
    if comparestr(term,'上学期') = 0 then
    begin
      delStu(6);
      delStu(9);
    end;

  end;

  procedure Database.setBeginYear(year : integer);
  var
    sql : string;
    par : Tb_parameters;
  begin
    //设置考学年份
    sql := 'update tb_parameters set '+
        ' averWenBase = 0,' +
        ' averWenTop  = 0,' +
        ' averLiBase  = 0,' +
        ' averLiTop   = 0,' +
        ' dev         = 0,' +
        ' greatRatio  = 0,' +
        ' yearBegin   = ' + inttostr(year) +
        ' where grade = 0';
    ado.ExecSqlStr(sql);
    par := Tb_parameters.Create;
    par.getPar();
  end;

  function Database.copyDatabase(year:integer; term:string; databaseModel : string):boolean;
  var
    databaseName : string;
    softPath : string;
    path : string;
    mes : string;  //对话框提示信息
    flag : boolean;
    sql : string;
  begin
    flag := true;
    //判断是否存在同名数据库
    softPath := ExtractFileDir(ParamStr(0));      //获取软件本身所在路径，用于文件操作
    path := softPath + '\database\' + inttostr(year)+'-'+inttostr(year+1)+'-'+term;
    if (FileExists(path+'.accdb')) or (FileExists(path+'.mdb')) then
    begin
      mes := inttostr(year)+'-'+inttostr(year+1)+'学年数据库已存在，是否删除原有数据库创建新数据库？';
      if MessageBox(0, PWideChar(mes), '操作选择', MB_OKCANCEL + MB_ICONQUESTION) = ID_OK then
      begin
        flag := true;   //是（覆盖）
      end else begin
        flag := false;
        result := false;
      end;
    end;

    if flag then
    begin
      databaseName := softPath +'\database\'+inttostr(year)+'-'+inttostr(year+1)+'-'+term+'.accdb';
      try
        CopyFile(PChar(databaseModel), PChar(databaseName), false);      //false表示替换同名旧文件
        result := true;
      except
        result := false;
      end;
    end;
  end;

  procedure Database.delCourse();
  var
    sql : string;
    course : Tb_course;
    I: Integer;
  begin
    course := Tb_course.Create;

    for I := 1 to 9 do
    begin

      sql := 'SELECT courseName,courseType from tb_course_'+inttostr(i);
      ado.SelectInfo(ADOQuery, sql);
      while not ADOQuery.eof do
      begin
        course.grade := i;
        course.courseName := ADOQuery.FieldByName('courseName').AsString;
        course.courseType := ADOQuery.FieldByName('courseType').AsString;
        course.delCourse;
        ADOQuery.Next;
      end;
    end;
  end;

  procedure Database.delPar();
  var
    sql : string;
  begin
    sql := 'update tb_parameters set '+
        ' averWenBase = -2,' +
        ' averWenTop  = 2,' +
        ' averLiBase  = -2,' +
        ' averLiTop   = 2,' +
        ' averVeto    = 6,' +
        ' dev         = 2,' +
        ' greatRatio  = 5,' +
        ' greatWeight = 16,' +
        ' inferiorWeight  = 2.2';
    ado.ExecSqlStr(sql);
  end;

  procedure Database.delTeacher();
  var
    sql : string;
  begin
    sql := 'update tb_classes set teacher = '+#39+#39;
    ado.ExecSqlStr(sql);
  end;

  procedure Database.delUser();
  var
    sql : string;
  begin
    sql := 'DELETE * FROM tb_users';
    ado.ExecSqlStr(sql);
    sql := 'INSERT INTO tb_users(users, psw, userIdentity, power) '+
        ' Values(' + #39 + 'admin'    +#39 +','+
                     #39 + '123'      +#39 +','+
                     '0, '+                     //0表示身份为管理员
                     #39 + '111111111'+#39 +')';
    ado.ExecSqlStr(sql);
  end;

  procedure Database.delStu(grade : integer);
  var
    classes : Tb_classes;
  begin
    classes := Tb_classes.Create;
    classes.selectByGrade(ADOQuery, grade);
    while not ADOQuery.eof do
    begin
      classes.grade := grade;
      classes.classID := ADOQuery.FieldByName('班号').AsString;
      classes.delClass();
      ADOQuery.next;
    end;
  end;

  procedure Database.updateCourseTable();
  var
    sql : string;
    I: Integer;
  begin
    //更新coures表
    i := 9;
    while (i>0) do
    begin
      sql := 'SELECT * into tb_course_'+inttostr(i+1)+' FROM tb_course_'+inttostr(i);
      ado.ExecSqlStr(sql);
      sql := 'drop   table tb_course_'+inttostr(i);
      ado.ExecSqlStr(sql);

      i := i-1;
    end;
    sql := 'drop   table tb_course_10';
      ado.ExecSqlStr(sql);
    sql := 'drop   table tb_course_7';
      ado.ExecSqlStr(sql);
    sql := 'Create TABLE tb_course_1('+
        'courseName VarChar(30) primary key,'+
        'courseType VarChar(10) not null)';
    ado.ExecSqlStr(sql);
    sql := 'Create TABLE tb_course_7('+
        'courseName VarChar(30) primary key,'+
        'courseType VarChar(10) not null)';
    ado.ExecSqlStr(sql);
  end;


  procedure Database.delExam(keepStu:boolean);
  var
    sql : string;
    exam : Tb_exam;
  begin
    exam := Tb_exam.Create;
    sql := 'SELECT examName,examGrade from tb_exam';
    ado.SelectInfo(ADOQuery, sql);
    while not ADOQuery.eof do
    begin
      exam.examName := ADOQuery.FieldByName('examName').AsString;
      exam.examGrade := ADOQuery.FieldByName('examGrade').AsString;
      exam.delExam(keepStu);
      ADOQuery.Next;
    end;

    exam.examName := '期末';
    exam.examGrade := '111111111';
    exam.addExam();
  end;

  procedure Database.saveConnString();
  begin
    connString := ADOOperate.connStr;
  end;

end.
