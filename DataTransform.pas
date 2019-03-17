unit DataTransform;

interface
uses
    Windows, Messages, Dialogs, SysUtils, StrUtils, GlobalData;
type

  CTransform = class
    Constructor Create;
  private
  public
    function classIDToGrade(classID : string):integer;
    function GradeToClassID(grade : integer):integer;
    function classIDToText(classID : string):string;
    function numToText(num : integer):string;
    function TextToNum(Text : string):integer;
end;

implementation
  Constructor CTransform.Create;
  begin
  end;

  {
      根据班号计算当前所处年级
      参数：  classID     班号
      返回值：年级
  }
  function CTransform.classIDToGrade(classID : string):integer;
  var
    y : string;
    temp : integer;
  begin
    y := LeftStr(classID,2);   //classID编码规范：例如1701表示17年入学的01班
    try
      temp := strtoint(y);
      result := (GlobalData.year mod 100) - temp + 1;
    except
      showmessage('从班号计算年级错误：班号不符合编码规范');
    end;
  end;

  {
      根据年级计算入学年份
      参数：  grade     年级
      返回值：入学年份（例如：2017年返回17）
  }
  function CTransform.GradeToClassID(grade : integer):integer;
  begin
    result := (GlobalData.year mod 100) - grade + 1;
  end;


  function CTransform.classIDToText(classID : string):string;
  var
    grade : integer;
    text : string;
    classID_3 : integer;
    classID_4 : integer;
  begin
     if strtoint(classID) < 1000 then
    begin
      text := numToText(strtoint(MidStr(classID, 1, 1))) + '年级';
      result := text;
      exit;
    end else begin
      grade := classIDToGrade(classID);
      classID_3 := strtoint(MidStr(classID, 3, 1));
      classID_4 := strtoint(MidStr(classID, 4, 1));

      text := numToText(grade);

      if classID_3 = 0  then
      begin
        text := text + '(' + numToText(classID_4) + ')';
      end else if classID_3 = 1 then begin
        text := text + '(十' + numToText(classID_4) + ')';
      end else begin
        text := text + '(' + numToText(classID_3) +'十' + numToText(classID_4) + ')';
      end;
      result := text;
    end;
  end;

  function CTransform.numToText(num : integer):string;
  begin
    case num of
    0: result := '';
    1: result := '一';
    2: result := '二';
    3: result := '三';
    4: result := '四';
    5: result := '五';
    6: result := '六';
    7: result := '七';
    8: result := '八';
    9: result := '九';
    10: result := '十';
    end;
  end;



  function CTransform.TextToNum(Text : string):integer;
  begin
    if comparestr('一',Text) = 0 then result := 1;
    if comparestr('二',Text) = 0 then result := 2;
    if comparestr('三',Text) = 0 then result := 3;
    if comparestr('四',Text) = 0 then result := 4;
    if comparestr('五',Text) = 0 then result := 5;
    if comparestr('六',Text) = 0 then result := 6;
    if comparestr('七',Text) = 0 then result := 7;
    if comparestr('八',Text) = 0 then result := 8;
    if comparestr('九',Text) = 0 then result := 9;
  end;
end.
