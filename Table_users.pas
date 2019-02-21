unit Table_users;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate;

type
    Tb_users = class
    Constructor Create;
  private
    ado : TAdoOperate;      //���ݿ��������
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
      ��֤�û���½��Ϣ
      ������userName    �û���
            password    ����
      ����ֵ���û���ݱ�ʶ��0��ʾ����Ա��1��ʾ�ɼ�¼����Ա
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
      �޸Ĺ���Ա�˺�����
      ������  userName    �û���
              oldPsw      ������
              newPsw      ������
      ����ֵ  0           �����벻��ȷ
              1           �޸ĳɹ�
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
    sql := 'SELECT users as �û���,'+
        ' psw as ����,'+
        ' power as Ȩ��'+
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
                     '1, '+       //1��ʾ���Ϊ�ɼ�¼����Ա
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
