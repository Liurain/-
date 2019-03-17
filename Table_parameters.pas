unit Table_parameters;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, GlobalData;

type
    Tb_parameters = class
    Constructor Create;
  private
    ado : TAdoOperate;        //���ݿ��������
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
    sql := 'SELECT grade       as �꼶, '          +
                 ' averWenBase as �Ŀ�ƽ�����½�, '+
                 ' averWenTop  as �Ŀ�ƽ�����Ͻ�, '+
                 ' averLiBase  as ���ƽ�����½�, '+
                 ' averLiTop   as ���ƽ�����Ͻ�, '+
                 ' averWenNo   as �Ŀƾ��ַ��, '+
                 ' averWenYes  as �Ŀƾ��ֿ϶�, '+
                 ' averLiNo    as ��ƾ��ַ��, '+
                 ' averLiYes   as ��ƾ��ֿ϶�, '+
                 ' dev         as ��׼��������ֵ, '+
                 ' greatRatio  as ������������ֵ, '+
                 ' greatWeight as ����ռ��, '+
                 ' inferiorWeight  as ѧ��ռ�� '+
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
