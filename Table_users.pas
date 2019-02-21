unit Table_users;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate;

type
    Tb_users = class
    Constructor Create;
  private
    ado : TAdoOperate;      //数据库操作对象
    ADOQuery: TADOQuery;
  public
    userName : string;
    password : string;
    power : string;


    function checkLogin(userName, password : string; var power : string):integer;
    function changeAdminInfo(userName, oldPsw, newPsw : string):integer;

    procedure selectUser(var ADOQuery : TADOQuery);
    procedure addUser(var ADOQuery : TADOQuery);
    procedure changeUser(var ADOQuery : TADOQuery);
    procedure delUser(var ADOQuery : TADOQuery);
  end;

implementation
  Constructor Tb_users.Create;
  begin
    ado := TAdoOperate.Create;
    ADOQuery := TADOQuery.Create(nil);
  end;

  {
      验证用户登陆信息
      参数：userName    用户名
            password    密码
      返回值：用户身份标识，0表示管理员，1表示成绩录入人员
  }
  function Tb_users.checkLogin(userName, password : string; var power : string):integer;
  var
    sql :string;
  begin
    result := -1;
    sql := 'select userIdentity, power from tb_users '+
        ' where users ='+#39+userName+#39+' and psw = '+#39+password+#39;
    ado.SelectInfo(ADOQuery,sql);

    power := ADOQuery.FieldByName('power').AsString;
    try
      if ADOQuery.RecordCount <> 0 then
        result := ADOQuery.FieldByName('userIdentity').AsInteger;
    except
      result := -1;
    end;
  end;

  {
      修改管理员账号密码
      参数：  userName    用户名
              oldPsw      旧密码
              newPsw      新密码
      返回值  0           旧密码不正确
              1           修改成功
  }
  function Tb_users.changeAdminInfo(userName, oldPsw, newPsw : string):integer;
  var
    sql : string;
    identity : integer;
  begin
    identity := -1;
    result := -1;
    sql := 'select userIdentity from tb_users '+
        ' where users ='+#39+userName+#39+' and psw = '+#39+oldPsw+#39;
    ado.SelectInfo(ADOQuery,sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      identity := ADOQuery.FieldByName('userIdentity').AsInteger;
    end else begin
      result := 0;
    end;

    if identity = 0 then
    begin
      sql := 'update tb_users set '+
        ' psw = ' + #39 + newPsw + #39 +
        ' where users=' + #39 + userName + #39;
      ado.ExecSqlStr(ADOQuery, sql);
      result := 1;
    end;
  end;

  procedure Tb_users.selectUser(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT users as 用户名,'+
        ' psw as 密码,'+
        ' power as 权限'+
        ' from tb_users'+
        ' order by users asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  procedure Tb_users.addUser(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'INSERT INTO tb_users(users, psw, userIdentity, power) '+
        ' Values(' + #39 + userName +#39 +','+
                     #39 + password +#39 +','+
                     '1, '+       //1表示身份为成绩录入人员
                     #39 + power+#39 +')';

    ado.ExecSqlStr(ADOQuery, sql);
  end;

  procedure Tb_users.changeUser(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'update tb_users set '+
        ' psw= ' + #39 + password +#39 +','+
        ' power=' + #39 + power+#39 +
        ' where userName=' + #39 + userName + #39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;

  procedure Tb_users.delUser(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'DELETE FROM tb_users' +
        ' WHERE users ='+#39 + userName +#39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;

end.
