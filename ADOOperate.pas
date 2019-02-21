unit ADOOperate;

interface

uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB;
type
    TAdoOperate = class
    Constructor Create;
  private
    databaseUser : string;
    dataPsw : string;
    connStr1 : string;
  public
    procedure openConn();
    procedure closeConn();
    procedure SelectInfo(var ADOQuery:TADOQuery; sqlStr: string);
    function ExecSqlStr(var ADOQuery:TADOQuery;sqlStr : string): integer;
//    procedure importBegin();
//    function importData(var ADOQuery:TADOQuery;sqlStr : string): integer;
//    procedure importEnd();
    function testConn() : boolean;
    procedure initConnStr(path : string);
    procedure setConnStr(path : string);     //这个使用私有连接串
end;
var
  connStr : string;         //全局变量，数据库连接串
  ADOConn: TADOConnection;


implementation

Constructor TAdoOperate.create;
begin
  ADOConn := TADOConnection.Create(nil);
  connStr1 := '';
end;

procedure TAdoOperate.initConnStr(path : string);
begin
  connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;'+
        'Data Source=' + path + ';'+
        'Persist Security Info=False;'+
        'Jet OLEDB:Database Password=123456';
end;

procedure TAdoOperate.setConnStr(path : string);
begin
  connStr1 := 'Provider=Microsoft.ACE.OLEDB.12.0;'+
        'Data Source=' + path + ';'+
        'Persist Security Info=False'+
        'Jet OLEDB:Database Password=123456';
end;

function TAdoOperate.testConn() : boolean;
begin
  openConn();
  if ADOConn.Connected then
  begin
    result := true;
    ADOConn.Close;
  end else begin
    result := false;
  end;
end;


procedure TAdoOperate.openConn();
var
  path : String;
begin
// 联接access数据库
  if ADOConn.Connected  then  ADOConn.Close;
  if Comparestr('',connStr1) = 0 then     //私有连接串为空时使用公有连接串
  begin
    ADOConn.ConnectionString := connStr;
  end else begin
    ADOConn.ConnectionString := connStr1;
  end;

  ADOConn.KeepConnection := true;
  ADOConn.LoginPrompt := false;
  ADOConn.ConnectOptions := coAsyncConnect; //must be coAsyncConnect!
  try
    ADOConn.Open();
    if(ADOConn.Connected = false)
    then
      showMessage('数据库连接失败');
  except
  end;
end;

procedure TAdoOperate.closeConn();
begin
  ADOConn.Close;
end;


procedure TAdoOperate.SelectInfo(var ADOQuery:TADOQuery; sqlStr: string);
begin
  openConn();
  try
    ADOQuery.Connection := ADOConn; //意思是ADOQuery连接数据库时用ADOConnection建立好的连接。
    ADOQuery.SQL.Text := sqlStr;
    ADOQuery.Open; //执行查询，如果是增、删、改则用pADOQ.ExecSQL
  finally
    ADOQuery.Close;
    closeConn();
    ADOQuery.Active := true;
  end;
end;



function TAdoOperate.ExecSqlStr(var ADOQuery:TADOQuery; sqlStr : String): integer;
begin
  openConn();
  try
    result := 0;
    ADOQuery.Connection := ADOConn;
    ADOQuery.Close;
    ADOQuery.SQL.Clear;
    ADOQuery.SQL.Add(sqlStr);
    result := ADOQuery.ExecSQL;  ////执行增、删、改
  finally
    ADOQuery.Close;
    closeConn();
  end;
end;

end.
