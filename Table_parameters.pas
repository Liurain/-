unit Table_parameters;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, GlobalData;

type
    Tb_parameters = class
    Constructor Create;
  private
    ado : TAdoOperate;        //数据库操作对象
  public
    procedure getPar();
    procedure getGradePar(var ADOQuery : TADOQuery);
    procedure changePar(AWB, AWT, ALB, ALT, averWenNo,averWenYes,averLiNo,averLiYes, dev1, greatRatio1 ,GW,IW: real; grade : integer);
  end;
implementation
  Constructor Tb_parameters.create;
  begin
    ado := TAdoOperate.Create;

  end;

  procedure Tb_parameters.getPar();
  var
    sql : string;
    ADOQuery : TADOQuery;
  begin
    ADOQuery := TADOQuery.Create(nil);
    sql := 'SELECT averWenBase, averWenTop, averLiBase, averLiTop, dev, greatRatio, yearBegin, averWenNo,averWenYes,averLiNo,averLiYes '+
        ' from tb_parameters where grade = 0';
    ado.SelectInfo(ADOQuery, sql);

    GlobalData.averWenBase := ADOQuery.FieldByName('averWenBase').AsFloat;
    GlobalData.averWenTop := ADOQuery.FieldByName('averWenTop').AsFloat;
    GlobalData.averLiBase := ADOQuery.FieldByName('averLiBase').AsFloat;
    GlobalData.averLiTop := ADOQuery.FieldByName('averLiTop').AsFloat;
    GlobalData.dev := ADOQuery.FieldByName('dev').AsFloat;
    GlobalData.greatRatio := ADOQuery.FieldByName('greatRatio').AsFloat;
    GlobalData.year := ADOQuery.FieldByName('yearBegin').AsInteger;
  end;

  procedure Tb_parameters.getGradePar(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT grade       as 年级, '          +
                 ' averWenBase as 文科平均分下界, '+
                 ' averWenTop  as 文科平均分上界, '+
                 ' averLiBase  as 理科平均分下界, '+
                 ' averLiTop   as 理科平均分上界, '+
                 ' averWenNo   as 文科均分否决, '+
                 ' averWenYes  as 文科均分肯定, '+
                 ' averLiNo    as 理科均分否决, '+
                 ' averLiYes   as 理科均分肯定, '+
                 ' dev         as 标准差评估差值, '+
                 ' greatRatio  as 优生率评估差值, '+
                 ' greatWeight as 优生占比, '+
                 ' inferiorWeight  as 学困占比 '+
        ' from tb_parameters where grade <> 0';
    ado.SelectInfo(ADOQuery, sql);
  end;

  procedure Tb_parameters.changePar(AWB, AWT, ALB, ALT,averWenNo,averWenYes,averLiNo,averLiYes, dev1, greatRatio1,GW,IW : real; grade : integer);
  var
    sql :string;
  begin
    sql := 'update tb_parameters set '+
        ' averWenBase = ' + floattostr(AWB)             + ',' +
        ' averWenTop  = ' + floattostr(AWT)             + ',' +
        ' averLiBase  = ' + floattostr(ALB)             + ',' +
        ' averLiTop   = ' + floattostr(ALT)             + ',' +
        ' averWenNo   = ' + floattostr(averWenNo)       + ',' +
        ' averWenYes  = ' + floattostr(averWenYes)      + ',' +
        ' averLiNo    = ' + floattostr(averLiNo)        + ',' +
        ' averLiYes   = ' + floattostr(averLiYes)       + ',' +
        ' dev         = ' + floattostr(dev1)            + ',' +
        ' greatRatio  = ' + floattostr(greatRatio1)     + ',' +
        ' greatWeight = ' + floattostr(GW)              + ',' +
        ' inferiorWeight  = ' + floattostr(IW)          +
        ' where grade = '+inttostr(grade);
    ado.ExecSqlStr(sql);
  end;
end.
