unit Table;

interface
uses
    Classes;
type
{*********************************�༶��ʵ��*********************************}
{*************************��Ӧ���ݿ��tb_classes***************************}
  tb_classes = record
    classID : string;
    stuNum : integer;
    teacher : string;
  end;

{******************************�༶�������ʵ��******************************}
{***********************��Ӧ���ݿ��tb_class_���**************************}
  tb_class = record
    stuID : string;
    stuName : string;
    stuSex : string;
    stuAge : integer;
  end;


{*******************************�꼶�γ̱�ʵ��********************************}
{***********************��Ӧ���ݿ��tb_course_�꼶**************************}
  teacherOfACourse = record
    TName : string;
    classID : string;
  end;
  tb_course = record
    courseName : string;
    teachers : array of teacherOfACourse;
  end;

{*****************************�꼶�γ̳ɼ���ʵ��******************************}
{*********************��Ӧ���ݿ��tb_scores_�꼶_�γ�***********************}
  examScore = record
    examName : string;              //��������(���磺��ѧ����ĩ)
    score : double;                 //���Է���
  end;
  tb_scores = record
    stuID : string;
    classID : string;
    stuName : string;
    scores : array of examScore;
  end;

{*********************************������ʵ��*********************************}
{************************��Ӧ���ݿ��tb_parameters*************************}
  tb_parameters = record
    ID : integer;
    maxScore : integer;             //�����Ծ�����
    year : integer;                 //��ѧ���
  end;

{*********************************���Ա�ʵ��*********************************}
{**************************��Ӧ���ݿ��tb_rexam****************************}
  tb_exam = record
    examName : string;
    examDate : TDateTime;
  end;

{*******************************���������ʵ��*******************************}
{**********************��Ӧ���ݿ��tb_report_�γ���************************}
  tb_report = record
    classID : string;               //���
    testNum : integer;              //ʵ�ʿ�������
    average : double;               //ƽ����
    standardDev :double;            //��׼��
    greatStuNum : integer;          //����ѧ������
    greatRatio :double;             //������
    inferiorStuNum : integer;       //ѧ��������
    evaluate : string;              //�춯(�Խ�ʦ������)
    teacher : string;               //�ον�ʦ
  end;

implementation

end.
