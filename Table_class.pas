unit Table_class;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, DataTransform;

type
    Tb_class = class
    Constructor Create;
  private
    ado : TAdoOperate;      //���ݿ��������
    procedure changeScoreTable(flag : integer);
  public
    classID : string;       //���
    stuID :string;          //ѧ��
    stuName : string;       //ѧ������
    sex : string;           //ѧ���Ա�
    age : integer;          //ѧ������

    procedure getStu(var ADOQuery : TADOQuery);
    procedure addStu();
    procedure changeStu();
    procedure delStu();
  end;

implementation
  Constructor Tb_class.create;
  begin
    ado := TAdoOperate.Create;
  end;

  {
      ���ݰ༶��ȡѧ����Ϣ
      �����Ϣ��
          ѧ��    stuID
          ����    stuName
          �Ա�    stuSex
          ����    stuAge
      Ԥ���趨������
          classID    ���
  }
  procedure Tb_class.getStu(var ADOQuery : TADOQuery);
  var
    sql :string;
  begin
    sql := 'SELECT stuID as ѧ��, '+
        ' stuName as ����, '+
        ' stuSex  as �Ա�, '+
        ' stuAge  as ����  '+
        ' from tb_class_'+classID+
        ' order by stuID asc';
    ado.SelectInfo(ADOQuery, sql);
  end;

  {
      ���һ��ѧ����Ϣ
      Ԥ���趨������
          classID    ���
          stuID      ѧ��
          stuName    ����
          sex        �Ա�
          age        ����
  }
  procedure Tb_class.addStu();
  var
    sql :string;
    ADOQuery : TADOQuery;
  begin
    ADOQuery := TADOQuery.Create(nil);
    //�ж����ݿ����Ƿ���ڸ�ѧ��
    sql := 'select * from tb_class_'+classID+' where stuID = '+#39+stuID+#39;
    ado.SelectInfo(ADOQuery, sql);
    if ADOQuery.RecordCount <> 0 then
    begin
      showmessage('���ݿ����Ѵ���ѧ����'+stuID+'   '+stuName);
      exit;
    end;
    //��ѧ��������������һ����¼
    sql := 'INSERT INTO tb_class_'+classID+'(stuID,stuName,stuSex,stuAge) Values(' +
          #39 + stuID +#39 + ',' +
          #39 + stuName + #39 + ',' +
          #39 + sex + #39 +','+
          #39 + inttostr(age) + #39 +')';
    ado.ExecSqlStr(sql);
    //�ð༶ѧ��������һ
    sql := 'update tb_classes set stuNum=stuNum + 1 where classID='+#39+classID+#39;
    ado.ExecSqlStr(sql);
    //��ѧ�������ڿγ̵ĳɼ���Ҳ����һ����¼
    changeScoreTable(0);  //0��ʾ����һ����¼
  end;


  {
      �޸�һ��ѧ����Ϣ
      Ԥ���趨������
          classID    ���
          stuID      ѧ��
          stuName    ����
          sex        �Ա�
          age        ����
  }
  procedure Tb_class.changeStu();
  var
    sql :string;
  begin
    sql := 'update tb_class_' + classID + ' set '+
        ' stuID   = ' + #39 + stuID   + #39 + ',' +
        ' stuName = ' + #39 + stuName + #39 + ',' +
        ' stuSex  = ' + #39 + sex     + #39 + ',' +
        ' stuAge  = ' + inttostr(age) +
        ' where stuID=' + #39 + stuID + #39;
    ado.ExecSqlStr(sql);
  end;


  {
      ɾ��һ��ѧ����Ϣ
      Ԥ���趨������
          stuID       ѧ��
  }
  procedure Tb_class.delStu();
  var
    sql :string;
  begin
    sql := 'DELETE FROM tb_class_' + classID +
        ' WHERE stuID = '+#39+stuID+#39;
    ado.ExecSqlStr(sql);
    //�ð༶ѧ��������һ
    sql := 'update tb_classes set stuNum=stuNum - 1 where classID='+#39+classID+#39;
    ado.ExecSqlStr(sql);

    //��ѧ�������ڿγ̵Ŀγ̱�Ҳɾ��һ����¼
    changeScoreTable(1);  //1��ʾɾ��һ����¼
  end;

  procedure Tb_class.changeScoreTable(flag : integer);
  var
    sql : string;
    courseList : TStringList;
    transform : CTransform;
    grade : integer;
    i : integer;
    ADOQuery : TADOQuery;
  begin
    ADOQuery := TADOQuery.Create(nil);
    courseList := TStringList.Create;
    transform := CTransform.Create;
    grade := transform.classIDToGrade(classID);
    sql := 'select courseName from tb_course_'+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
    while not ADOQuery.eof do
    begin
      courseList.Add(ADOQuery.FieldByName('courseName').AsString);
      ADOQuery.Next;
    end;

    for I := 0 to courseList.Count - 1 do
    begin
      case flag of
        0:begin      //����һ����¼
          sql := 'INSERT INTO tb_scores_'+inttostr(grade)+'_'+courseList[i]+'(stuID,classID,stuName) Values(' +
                    #39 + stuID   + #39 + ',' +
                    #39 + classID + #39 + ',' +
                    #39 + stuName + #39 + ')';
        end;
        1:begin      //ɾ��һ����¼
          sql := 'DELETE FROM tb_scores_'+inttostr(grade)+'_'+courseList[i]+
                    ' WHERE stuID = '+#39+stuID+#39;
        end;
      end;

      ado.ExecSqlStr(sql);
    end;
  end;

end.
