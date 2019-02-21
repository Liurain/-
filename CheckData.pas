unit CheckData;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate;

type
    TCheckData = class
    Constructor Create;

  private
    maxScore : real;
  public
    procedure setMaxScore(Score : real);

  {
      所有检查都允许参数为默认值(空/0）
      当参数为默认值时合法，flag返回为true
  }
    function grage(grade : integer; var flag:boolean) : integer;
    function classID(classID : string; var flag:boolean) : string;
    function stuNum(StuNum : string; var flag:boolean) : integer;
    function stuID(StuID : string; var flag:boolean) : string;
    function stuSex(StuSex : integer; var flag:boolean) : string;
    function stuAge(StuAge : string; var flag:boolean) : integer;
    function score(score : string; var flag:boolean) : real;
    function Par(Par : string; var flag:boolean) : real;
    function term(term : integer; var flag:boolean) : string;
    function courseType(courseType : string; var flag:boolean) : string;
    function ratio(ratio : string; var flag:boolean) : real;
    function year(year : string; var flag:boolean) : integer;
    function courseName(courseName : string; var flag:boolean):string;
    function stuName(stuName : string; var flag:boolean):string;
  end;

implementation
  Constructor TCheckData.create;
  begin
    maxScore := 100;  //初始默认满分为100
  end;

  procedure TCheckData.setMaxScore(Score : real);
  begin
    maxScore := Score;
  end;

  function TCheckData.grage(grade : integer; var flag:boolean) : integer;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    grade := grade + 1;  //ComboBox编号从0开始
    if flag then
    begin
      if grade = 0 then
      begin
        showmessage('请选择添需要加班级所在的年级');
        flag := false;
      end;
    end;
    result := grade;
  end;


  function TCheckData.classID(classID : string; var flag:boolean) : string;
  var
    ID : integer;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if classID = '' then
    begin
      //允许班级为默认值（空）
      result := classID;
      exit;
    end else begin
      //判断班级格式是否正确
      try
        ID := strtoint(classID);
        if (ID<0) or (ID>999999) then
        begin
          showMessage('请规范填写班号(例如17年入学01班为1701)');
          flag := false;
        end else begin
          result := classID;
        end;
      except
        showMessage('班号必须是一个整数(例如17年入学01班为1701)');
        flag := false;
      end;
    end;
  end;


  function TCheckData.stuNum(StuNum : string; var flag:boolean) : integer;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if StuNum = '' then
    begin
      //允许班级人数为默认值（空）
      result := 0;
      exit;
    end else begin
      try
        result := strtoint(stuNum);
      except
        showMessage('学生人数请输入一个整数');
        flag := false;
      end;
    end;
  end;


  function TCheckData.stuID(StuID : string; var flag:boolean) : string;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if StuID = '' then
    begin
      //学生学号不可以为空
      showmessage('本操作学号不可以为空');
      flag := false;
      exit;
    end else begin
      try
        strtoint(stuID);   //默认学号为一串纯数字
        result :=  stuID;
      except
        showMessage('请输入符合规范的学号（纯数字）');
        flag := false;
      end;
    end;
  end;


  function TCheckData.stuSex(StuSex : integer; var flag:boolean) : string;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    case stuSex of
      0: result := '男';
      1: result := '女';
      else
        flag := false;
    end;

  end;


  function TCheckData.stuAge(StuAge : string; var flag:boolean) : integer;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if StuAge = '' then
    begin
      //允许班级人数为默认值（空）
      result := 0;
      exit;
    end else begin
      try
        result := strtoint(StuAge);
      except
        showMessage('年龄请输入一个数字');
        flag := false;
      end;
    end;
  end;

  function TCheckData.score(score : string; var flag:boolean) : real;
  var
    s : real;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if score = '' then
    begin
      //允许班级人数为默认值（空）
      result := 0;
      exit;
    end;


    try
      s := strtofloat(Score);
    except
      showMessage('请输入一个数字');
      flag := false;
      exit;
    end;

    if (s<0) or (s>maxScore)  then
    begin
      showmessage('合法的参数应该在0-'+floattostr(maxScore)+'之间');
      flag := false;
      exit;
    end;

    result := s;
  end;

  function TCheckData.par(par : string; var flag:boolean) : real;
  var
    s : real;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if par = '' then
    begin
      showmessage('参数不可以有空值');
      flag := false;
      exit;
    end;


    try
      s := strtofloat(par);
    except
      showMessage('请输入一个数字');
      flag := false;
      exit;
    end;

    if (s<(-100)) or (s>100)  then
    begin
      showmessage('合法的参数应该在-100~'+floattostr(maxScore)+'之间');
      flag := false;
      exit;
    end;

    result := s;
  end;

  function TCheckData.term(term : integer; var flag:boolean) : string;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    case term of
      -1: begin //空值
        flag := false;
      end;
      0: begin  //上学期
        result := '上';
      end;
      1: begin  //下学期
        result := '下';
      end;
    end;
  end;

  function TCheckData.courseType(courseType : string; var flag:boolean) : string;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if Comparestr(courseType,'文科') = 0 then
    begin
      result := '文科';
    end else if Comparestr(courseType,'理科') = 0 then
    begin
      result := '理科';
    end else begin
      flag := false;
    end;
  end;

  function TCheckData.ratio(ratio : string; var flag:boolean) : real;
  var
    r : real;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if Comparestr(ratio,'') = 0 then
    begin
      showmessage('优生率参不可以为空');
      result := 0;
      exit;
    end;


    try
      r := strtofloat(ratio);
    except
      showMessage('优生率参数应该是0~100的数字');
      flag := false;
      exit;
    end;

    if (r<0) or (r>100)  then
    begin
      showMessage('优生率参数应该是0~100的数字');
      flag := false;
      exit;
    end;

    result := r;
  end;

  function TCheckData.year(year: string; var flag: Boolean):integer;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if Comparestr(year,'') = 0 then
    begin
      showmessage('年份不可为空！');
      flag := false;
      exit;
    end else begin
      try
        result := strtoint(year);
      except
        showMessage('年份请输入一个数字');
        flag := false;
      end;
    end;
  end;

  function TCheckData.courseName(courseName : string; var flag:boolean):string;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if Comparestr(courseName ,'') = 0 then
    begin
      showmessage('课程名不可以为空。');
    end else begin
      result := courseName;
    end;

  end;

  function TCheckData.stuName(stuName : string; var flag:boolean):string;
  begin
    if not flag then exit; //若flag为false则不执行后续判断

    if Comparestr(stuName ,'') = 0 then
    begin
      showmessage('学生姓名不可以为空。');
    end else begin
      result := stuName;
    end;

  end;
end.
