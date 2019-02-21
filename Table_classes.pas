unit Table_classes;

interface

uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, DataTransform;

type
    Tb_classes = class
    Constructor Create;
  private
    ado : TAdoOperate;
  public
    grade : integer;      //�꼶
    classID: string;      //���
    stuNum: integer;      //ѧ������
    teacher: string;      //������

    procedure selectByGrade(var ADOQuery : TADOQuery; grade1 : integer);
    procedure selectAll(var ADOQuery : TADOQuery);

    procedure addClass(var ADOQuery : TADOQuery);
    procedure changeClass(var ADOQuery : TADOQuery);
    procedure delClass(var ADOQuery : TADOQuery);
end;


implementation
  Constructor Tb_classes.create;
  begin
    ado := TAdoOperate.Create;
  end;

  procedure Tb_classes.selectAll(var ADOQuery : TADOQuery);
  var
    sql : string;
  begin
    sql := 'SELECT classID as ���, '+
        ' stuNum as ѧ������, '+
        ' teacher as ������ '+
        ' from tb_classes '+
        ' order by classID asc';
    ado.SelectInfo(ADOQuery, sql);
  end;


  {
      �����꼶��ȡ���꼶�����Ϣ
      ������grade --�꼶
  }
  procedure Tb_classes.selectByGrade(var ADOQuery : TADOQuery; grade1 : integer);
  var
    sql : string;
    year : string; //ѧ���е����
    transform : CTransform;
  begin
    transform := CTransform.Create;
    year := inttostr(transform.GradeToClassID(grade1));
    sql := 'SELECT classID as ���, '+
        ' stuNum as ѧ������, '+
        ' teacher as ������ '+
        ' from tb_classes '+
        ' where classID like '+ year +'& "%"'+
        ' order by classID asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      �����ݿ������һ���༶
      ������
            classID   --��ţ�����01
            stuNum    --���༶ѧ������
            teacher   --������
            grade     --�꼶
  }
  procedure Tb_classes.addClass(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    //�ж����ݿ����Ƿ���ڸİ༶
    sql := 'select * from tb_classes where classID = '+#39+classID+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('���ݿ����Ѵ���'+classID+'��');
      exit;
    end;

    //��tb_classes�������һ���༶��Ϣ
    if strtoint(classID) < 1000 then
    begin
      classID := '0'+ inttostr(strtoint(classID));
    end;
    sql := 'INSERT INTO tb_classes(classID,stuNum,teacher) Values(' +
            #39 + classID + #39 + ',' +
            #39 + inttostr(stuNum) + #39 +',' +
            #39 + teacher + #39 +')';
    ado.ExecSqlStr(ADOQuery, sql);

    //Ϊ�ð༶����һ�Ż������
    sql := 'Create TABLE tb_class_' + classID +'('+
        'stuID VarChar(10) primary key,'+
        'stuName VarChar(10) not null,'+
        'stuSex VarChar(5) default '+ #39 +'��'+ #39+','+
        'stuAge int default 0 )';
    ado.ExecSqlStr(ADOQuery, sql);

    //�ڶ�Ӧ�꼶�Ŀγ̱��м���һ��
    sql := 'alter table tb_course_'+inttostr(grade)+
        ' add T_'+classID+' VarChar(10)  null';
    ado.ExecSqlStr(ADOQuery, sql);
  end;


  {
      �޸����ݿ���һ���༶����Ϣ
      ������grade         --�༶�����꼶
            classID_str   --��ţ�����01
            stuNum_str    --���༶ѧ������
            teacher_str   --������
  }
  procedure Tb_classes.changeClass(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'update tb_classes set '+
        ' stuNum = ' + inttostr(stuNum) + ',' +
        ' teacher =' + #39 + teacher + #39 +
        ' where classID=' + #39 + classID + #39;
    ado.ExecSqlStr(ADOQuery, sql);
  end;


  {
      ɾ�����ݿ���һ���༶����Ϣ
      ������grade         --�༶�����꼶
            classID_str   --��ţ�����01
            stuNum_str    --���༶ѧ������
            teacher_str   --������
  }
  procedure Tb_classes.delClass(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    //�Ӱ༶����ɾ���༶��Ϣ
    sql := 'DELETE FROM tb_classes WHERE classID = '+#39+classID+#39;
    ado.ExecSqlStr(ADOQuery, sql);
    //ɾ���ð༶��صı�
    sql := 'Drop table tb_class_'+classID;
    ado.ExecSqlStr(ADOQuery, sql);
    //ɾ���γ̱��������
    sql := 'ALTER TABLE tb_course_'+inttostr(grade)+' DROP COLUMN T_'+classID;
    ado.ExecSqlStr(ADOQuery, sql);
  end;

end.
