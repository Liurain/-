unit Table;

interface
uses
    Classes;
type
{*********************************班级表实体*********************************}
{*************************对应数据库表：tb_classes***************************}
  tb_classes = record
    classID : string;
    stuNum : integer;
    teacher : string;
  end;

{******************************班级花名册表实体******************************}
{***********************对应数据库表：tb_class_班号**************************}
  tb_class = record
    stuID : string;
    stuName : string;
    stuSex : string;
    stuAge : integer;
  end;


{*******************************年级课程表实体********************************}
{***********************对应数据库表：tb_course_年级**************************}
  teacherOfACourse = record
    TName : string;
    classID : string;
  end;
  tb_course = record
    courseName : string;
    teachers : array of teacherOfACourse;
  end;

{*****************************年级课程成绩表实体******************************}
{*********************对应数据库表：tb_scores_年级_课程***********************}
  examScore = record
    examName : string;              //测试名称(例如：上学期期末)
    score : double;                 //测试分数
  end;
  tb_scores = record
    stuID : string;
    classID : string;
    stuName : string;
    scores : array of examScore;
  end;

{*********************************参数表实体*********************************}
{************************对应数据库表：tb_parameters*************************}
  tb_parameters = record
    ID : integer;
    maxScore : integer;             //测试试卷满分
    year : integer;                 //入学年份
  end;

{*********************************测试表实体*********************************}
{**************************对应数据库表：tb_rexam****************************}
  tb_exam = record
    examName : string;
    examDate : TDateTime;
  end;

{*******************************分析报告表实体*******************************}
{**********************对应数据库表：tb_report_课程名************************}
  tb_report = record
    classID : string;               //班号
    testNum : integer;              //实际考试人数
    average : double;               //平均分
    standardDev :double;            //标准差
    greatStuNum : integer;          //优秀学生人数
    greatRatio :double;             //优生率
    inferiorStuNum : integer;       //学困生人数
    evaluate : string;              //异动(对教师的评价)
    teacher : string;               //任课教师
  end;

implementation

end.
