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
      ���м�鶼�������ΪĬ��ֵ(��/0��
      ������ΪĬ��ֵʱ�Ϸ���flag����Ϊtrue
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
    maxScore := 100;  //��ʼĬ������Ϊ100
  end;

  procedure TCheckData.setMaxScore(Score : real);
  begin
    maxScore := Score;
  end;

  function TCheckData.grage(grade : integer; var flag:boolean) : integer;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    grade := grade + 1;  //ComboBox��Ŵ�0��ʼ
    if flag then
    begin
      if grade = 0 then
      begin
        showmessage('��ѡ������Ҫ�Ӱ༶���ڵ��꼶');
        flag := false;
      end;
    end;
    result := grade;
  end;


  function TCheckData.classID(classID : string; var flag:boolean) : string;
  var
    ID : integer;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if classID = '' then
    begin
      //����༶ΪĬ��ֵ���գ�
      result := classID;
      exit;
    end else begin
      //�жϰ༶��ʽ�Ƿ���ȷ
      try
        ID := strtoint(classID);
        if (ID<0) or (ID>999999) then
        begin
          showMessage('��淶��д���(����17����ѧ01��Ϊ1701)');
          flag := false;
        end else begin
          result := classID;
        end;
      except
        showMessage('��ű�����һ������(����17����ѧ01��Ϊ1701)');
        flag := false;
      end;
    end;
  end;


  function TCheckData.stuNum(StuNum : string; var flag:boolean) : integer;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if StuNum = '' then
    begin
      //����༶����ΪĬ��ֵ���գ�
      result := 0;
      exit;
    end else begin
      try
        result := strtoint(stuNum);
      except
        showMessage('ѧ������������һ������');
        flag := false;
      end;
    end;
  end;


  function TCheckData.stuID(StuID : string; var flag:boolean) : string;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if StuID = '' then
    begin
      //ѧ��ѧ�Ų�����Ϊ��
      showmessage('������ѧ�Ų�����Ϊ��');
      flag := false;
      exit;
    end else begin
      try
        strtoint(stuID);   //Ĭ��ѧ��Ϊһ��������
        result :=  stuID;
      except
        showMessage('��������Ϲ淶��ѧ�ţ������֣�');
        flag := false;
      end;
    end;
  end;


  function TCheckData.stuSex(StuSex : integer; var flag:boolean) : string;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    case stuSex of
      0: result := '��';
      1: result := 'Ů';
      else
        flag := false;
    end;

  end;


  function TCheckData.stuAge(StuAge : string; var flag:boolean) : integer;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if StuAge = '' then
    begin
      //����༶����ΪĬ��ֵ���գ�
      result := 0;
      exit;
    end else begin
      try
        result := strtoint(StuAge);
      except
        showMessage('����������һ������');
        flag := false;
      end;
    end;
  end;

  function TCheckData.score(score : string; var flag:boolean) : real;
  var
    s : real;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if score = '' then
    begin
      //����༶����ΪĬ��ֵ���գ�
      result := 0;
      exit;
    end;


    try
      s := strtofloat(Score);
    except
      showMessage('������һ������');
      flag := false;
      exit;
    end;

    if (s<0) or (s>maxScore)  then
    begin
      showmessage('�Ϸ��Ĳ���Ӧ����0-'+floattostr(maxScore)+'֮��');
      flag := false;
      exit;
    end;

    result := s;
  end;

  function TCheckData.par(par : string; var flag:boolean) : real;
  var
    s : real;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if par = '' then
    begin
      showmessage('�����������п�ֵ');
      flag := false;
      exit;
    end;


    try
      s := strtofloat(par);
    except
      showMessage('������һ������');
      flag := false;
      exit;
    end;

    if (s<(-100)) or (s>100)  then
    begin
      showmessage('�Ϸ��Ĳ���Ӧ����-100~'+floattostr(maxScore)+'֮��');
      flag := false;
      exit;
    end;

    result := s;
  end;

  function TCheckData.term(term : integer; var flag:boolean) : string;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    case term of
      -1: begin //��ֵ
        flag := false;
      end;
      0: begin  //��ѧ��
        result := '��';
      end;
      1: begin  //��ѧ��
        result := '��';
      end;
    end;
  end;

  function TCheckData.courseType(courseType : string; var flag:boolean) : string;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if Comparestr(courseType,'�Ŀ�') = 0 then
    begin
      result := '�Ŀ�';
    end else if Comparestr(courseType,'���') = 0 then
    begin
      result := '���';
    end else begin
      flag := false;
    end;
  end;

  function TCheckData.ratio(ratio : string; var flag:boolean) : real;
  var
    r : real;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if Comparestr(ratio,'') = 0 then
    begin
      showmessage('�����ʲβ�����Ϊ��');
      result := 0;
      exit;
    end;


    try
      r := strtofloat(ratio);
    except
      showMessage('�����ʲ���Ӧ����0~100������');
      flag := false;
      exit;
    end;

    if (r<0) or (r>100)  then
    begin
      showMessage('�����ʲ���Ӧ����0~100������');
      flag := false;
      exit;
    end;

    result := r;
  end;

  function TCheckData.year(year: string; var flag: Boolean):integer;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if Comparestr(year,'') = 0 then
    begin
      showmessage('��ݲ���Ϊ�գ�');
      flag := false;
      exit;
    end else begin
      try
        result := strtoint(year);
      except
        showMessage('���������һ������');
        flag := false;
      end;
    end;
  end;

  function TCheckData.courseName(courseName : string; var flag:boolean):string;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if Comparestr(courseName ,'') = 0 then
    begin
      showmessage('�γ���������Ϊ�ա�');
    end else begin
      result := courseName;
    end;

  end;

  function TCheckData.stuName(stuName : string; var flag:boolean):string;
  begin
    if not flag then exit; //��flagΪfalse��ִ�к����ж�

    if Comparestr(stuName ,'') = 0 then
    begin
      showmessage('ѧ������������Ϊ�ա�');
    end else begin
      result := stuName;
    end;

  end;
end.
