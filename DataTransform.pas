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
end;

implementation
  Constructor CTransform.Create;
  begin
  end;

  {
      ���ݰ�ż��㵱ǰ�����꼶
      ������  classID     ���
      ����ֵ���꼶
  }
  function CTransform.classIDToGrade(classID : string):integer;
  var
    y : string;
    temp : integer;
  begin
    y := LeftStr(classID,2);   //classID����淶������1701��ʾ17����ѧ��01��
    try
      temp := strtoint(y);
      result := (GlobalData.year mod 100) - temp + 1;
    except
      showmessage('�Ӱ�ż����꼶���󣺰�Ų����ϱ���淶');
    end;
  end;

  {
      �����꼶������ѧ���
      ������  grade     �꼶
      ����ֵ����ѧ��ݣ����磺2017�귵��17��
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
      text := numToText(strtoint(MidStr(classID, 1, 1))) + '�꼶';
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
        text := text + '(ʮ' + numToText(classID_4) + ')';
      end else begin
        text := text + '(' + numToText(classID_3) +'ʮ' + numToText(classID_4) + ')';
      end;
      result := text;
    end;
  end;

  function CTransform.numToText(num : integer):string;
  begin
    case num of
    0: result := '';
    1: result := 'һ';
    2: result := '��';
    3: result := '��';
    4: result := '��';
    5: result := '��';
    6: result := '��';
    7: result := '��';
    8: result := '��';
    9: result := '��';
    10: result := 'ʮ';
    end;
  end;

end.
