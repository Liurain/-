unit ADOOperate;

interface

uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB;
type
    TAdoOperate = class
    Constructor Create;
  private
    ADOQuery1 : TADOQuery;
  public
    procedure openConn();
    procedure closeConn();
    procedure SelectInfo(var ADOQuery:TADOQuery; sqlStr: string);
    function ExecSqlStr(sqlStr : string): integer;
    function testConn() : boolean;
    procedure initConnStr(path : string);
    procedure setConnStr(str : string);

    procedure beginTrans();
    function addSqlToTrans(sql : string): boolean;
    function commitTrans():boolean;
end;
var
  ADOConn: TADOConnection;
  connStr : string;


implementation

Constructor TAdoOperate.create;
begin
  ADOQuery1 := TADOQuery.Create(nil);
end;

procedure TAdoOperate.initConnStr(path : string);
begin
  connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;'+
        'Data Source=' + path + ';'+
        'Persist Security Info=False;'+
        'Jet OLEDB:Database Password=123456';
  openConn();
end;

procedure TAdoOperate.setConnStr(str : string);
begin
  connStr := str;
  openConn();
end;

procedure TAdoOperate.beginTrans();
begin
  ADOConn.BeginTrans();
end;

function TAdoOperate.addSqlToTrans(sql : string): boolean;
begin
  try
    ADOConn.Execute(sql);
    result := true;
  except
    result := false;
  end;
end;

function TAdoOperate.commiTtrans():boolean;
begin
  try
    ADOConn.CommitTrans;
    result := true;
  except
    ADOConn.RollbackTrans;
    result := false;
  end;
end;

function TAdoOperate.testConn() : boolean;
begin
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

  ADOConn.ConnectionString := connStr;
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
  try
    ADOQuery.Connection := ADOConn; //意思是ADOQuery连接数据库时用ADOConnection建立好的连接。
    ADOQuery.SQL.Text := sqlStr;
    ADOQuery.Open; //执行查询，如果是增、删、改则用pADOQ.ExecSQL
  finally
//    ADOQuery.Close;
//    closeConn();
    ADOQuery.Active := true;
  end;
end;

function TAdoOperate.ExecSqlStr(sqlStr : String): integer;
begin
  try
    result := 0;
    ADOQuery1.Connection := ADOConn;
    ADOQuery1.Close;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add(sqlStr);
    result := ADOQuery1.ExecSQL;  ////执行增、删、改
  finally
//    ADOQuery1.Close;
//    closeConn();
  end;
end;

initialization

  ADOConn := TADOConnection.Create(nil);
end.
