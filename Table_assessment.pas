unit Table_assessment;

interface
uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ADOOperate, Table_classes, Math, GlobalData,DataTransform,Table_parameters;

type
  assessment = record   //结果分析表实体
    classID : string;                   //班号
    testNum : integer;                  //实际考试人数
    average : double;                   //平均分
    standardDev :double;                //标准差
    greatStuNum : integer;              //优秀学生人数
    greatRatio :double;                 //优生率
    inferiorStuNum : integer;           //学困生人数
    inferiorRatio :double;              //学困率
    evaluate : string;                  //异动(对教师的评价)
    teacher : string;                   //任课教师
    greatLine : double;                 //优生分数线
    inferiorLine : double;              //学困生分数线
  end;

  Tb_assessment = class
    Constructor Create;
  private
    ado : TAdoOperate;
    ADOQuery: TADOQuery;
    examName : string;                //考试名称
    courseName : string;              //课程名称

    procedure createReportTable(exam, course : string);
    procedure fillInTheReportTable(exam, course : string);
    procedure getAllclass(grade : integer; var classList:TStringList);
    procedure getGradedata(grade : integer; var grade_ass : assessment);
    procedure BubbleSort(var x:array of double);
    procedure getClassdata(grade : integer; classID : string;var class_ass, grade_ass : assessment);
    procedure computeDev(var score : array of double; var average, dev : double);
    function evaluate(grade : integer; class_ass, grade_ass : assessment):string;  //教师评估
    function evaluateAver(base, top, averNo, averYes: double; class_aver, grade_aver: Double):integer;  //平均分评估
    function evaluateDev(dev, class_dev, grade_dev: Double):integer;  //标准差评估
    function evaluateGrateRatio(greatRatio, class_grateRatio, grade_grateRatio: Double):integer;  //优生率评估
    procedure initAss(var ass : assessment);  //初始化assessment对象
    procedure insertDataToTable(grade:integer;var ass : assessment);  //把记录插入到表中

  public
    procedure createReport(exam, course : string);
end;

implementation
  Constructor Tb_assessment.Create;
  begin
    ado := TAdoOperate.Create;
    ADOQuery := TADOQuery.Create(nil);
  end;

  procedure Tb_assessment.createReport(exam, course : string);
  var
    sql : string;
  begin
    examName := exam;
    courseName := course;
    //测试当前课程的报告是否存在
    try
      //删除原有报告表
      sql := 'Drop table tb_report_' + exam + '_' + course;
      ado.ExecSqlStr(sql);
    except
    end;
   //创建新的报告表
    createReportTable(examName, courseName);

    //更新报告中的数据
    fillInTheReportTable(examName, courseName);
  end;

  procedure Tb_assessment.createReportTable(exam, course : string);
  var
    sql : string;
  begin
    sql := 'Create TABLE tb_report_' + exam + '_' + course+'('+
        ' classID             VarChar(5)  primary key,'+
        ' testNum             int         default 0,'+
        ' average             float       default 0,'+
        ' standardDev         float       default 0,'+
        ' greatStuNum         int         default 0,'+
        ' greatRatio          float       default 0,'+
        ' inferiorStuNum      int         default 0,'+
        ' inferiorRatio       float       default 0,'+
        ' evaluate            VarChar(2)  null,'+
        ' teacher             VarChar(10) null,'+
        ' greatLine           float       default 0,'+
        ' inferiorLine        float       default 0)';
    ado.ExecSqlStr(sql);
  end;

  procedure Tb_assessment.fillInTheReportTable(exam, course : string);
  var
    sql : string;
    grade : integer;
    classList : TStringList;
    i : integer;

    class_ass : assessment;    //分析表中一个班的数据实体
    grade_ass : assessment;    //分析表中一个年级的数据实体
  begin
    classList := TStringList.Create;

    //从数据库中找出拥有该门课程的年级
    for grade := 1 to 9 do
    begin
      sql := 'select count(*) from tb_course_'+inttostr(grade)+
            ' where courseName ='+#39+course+#39;
      ado.SelectInfo(ADOQuery, sql);

      if (ADOQuery.Fields[0].asInteger <>0) then   //该年级拥有这门课程
      begin
        initAss(grade_ass);
        getGradedata(grade, grade_ass);    //获取该年级的相关数据
        classList.Clear;
        getAllclass(grade, classList);     //获取这个年级所有的班级编号

        for I := 0 to classList.Count - 1 do
        begin
          initAss(class_ass);
          getClassdata(grade, classList[i], class_ass, grade_ass); //从其他表中获取数据初始化这两个实体
          insertDataToTable(grade, class_ass);
        end;
        if classList.Count <> 0 then
        begin
          insertDataToTable(grade, grade_ass);
        end;
      end;
    end;
  end;

  procedure Tb_assessment.getAllclass(grade : integer; var classList:TStringList);
  var
    sql : string;
    y : integer;   //学号中的年份
    trans : CTransform;
  begin
    trans := CTransform.create;
    y := trans.GradeToClassID(grade);
    sql := 'SELECT classID  from tb_classes '+
        ' where classID like '+ inttostr(y) +'& "%"'+
        ' order by classID asc';
    ado.SelectInfo(ADOQuery, sql);

    while not ADOQuery.eof do
    begin
      classList.Add(ADOQuery.FieldByName('classID').AsString);
      ADOQuery.Next;
    end;
  end;

  procedure Tb_assessment.getGradedata(grade : integer; var grade_ass : assessment);
  var
    sql : string;
    scoreArr : array of double;

    testNum : integer;           //测试人数
    quekaoNum : integer;         //缺考人数
    average : double;            //平均值
    dev : double;                //标准差
    greatStuNum : integer;       //优生人数
    inferiorStuNum : integer;    //学困生人数

    greatLine : double;          //优秀最低分，含临界
    inferiorLine : double;       //学困的最高分，不含临界
    greatWeight : double;        //优生占有率
    inferiorWeight : double;      //学困占有率
    greatIndex : integer;        //
    inferiorIndex : integer;      //

    i : integer;
  begin
    testNum := 0;
    quekaoNum := 0;
    average := 0;
    dev := 0;
    greatStuNum := 0;
    inferiorStuNum := 0;
  //从参数表中获取优生比例和学困比例
    sql := 'select greatWeight, inferiorWeight from tb_parameters  where grade ='+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
    greatWeight := ADOQuery.FieldByName('greatWeight').Asfloat;
    inferiorWeight := ADOQuery.FieldByName('inferiorWeight').Asfloat;

  //从数据库中获取这个年级该课程所有成绩
    sql := 'select '+examName+' from tb_scores_'+inttostr(grade)+'_'+courseName;
    ado.SelectInfo(ADOQuery, sql);
    setLength(scoreArr,ADOQuery.RecordCount);

    for I := 0 to ADOQuery.RecordCount-1 do
    begin
      scoreArr[i] := ADOQuery.FieldByName(examName).AsFloat;
      ADOQuery.Next;
    end;

    BubbleSort(scoreArr);

    //计算优生和学困分数线
    for I := 0 to Length(scoreArr)-1 do
    begin
      if (scoreArr[i] <> -1) then
      begin
        testNum := testNum+1;
      end else begin
        quekaoNum := quekaoNum + 1;
      end;
    end;

    if testNum <> 0 then
    begin
      greatIndex := (Length(scoreArr)-Round(testNum*(greatWeight/100)));        //优生占有率
      inferiorIndex := (Round(testNum*(inferiorWeight/100))+quekaoNum);      //学困占有率
      if inferiorIndex = -1 then inferiorIndex := 0;    //无人考试时inferiorIndex=-1；


      greatLine := scoreArr[greatIndex];      //优生线
      inferiorLine := scoreArr[inferiorIndex];        //学困线
    end else begin
      greatLine := 0;      //优生线
      inferiorLine := 0;        //学困线
    end;

    computeDev(scoreArr, average, dev);       //计算平均值和标准差

    //统计优生和学困人数
    for I := 0 to Length(scoreArr)-1 do
    begin
      if (scoreArr[i] <> -1) then
      begin
        if (scoreArr[i]>= greatLine) then
          greatStuNum := greatStuNum+1;
        if (scoreArr[i]<inferiorLine) then
          inferiorStuNum := inferiorStuNum+1;
      end;
    end;

    if (testNum <> 0) then
    begin
      grade_ass.greatRatio := greatStuNum / testNum;
      grade_ass.inferiorRatio := inferiorStuNum / testNum;
    end else begin
      grade_ass.greatRatio := 0;
      grade_ass.inferiorRatio := 0;
    end;
    grade_ass.classID := inttostr(grade*100);
    grade_ass.testNum := testNum;
    grade_ass.average := average;
    grade_ass.standardDev := dev;
    grade_ass.greatStuNum := greatStuNum;
    grade_ass.inferiorStuNum := inferiorStuNum;
    grade_ass.evaluate := '';
    grade_ass.teacher := '';
    grade_ass.greatLine := greatLine;
    grade_ass.inferiorLine := inferiorLine;
  end;

  procedure Tb_assessment.BubbleSort(var x:array of double);       //冒泡排序(大的往后)
  var
    i,j : integer;
    doubleTmp : double;
  begin
    for i:=0 to high(x) do
    begin
      for j:=0 to high(x)-1 do
      begin
        if x[j]>x[j+1] then
        begin
          doubleTmp:=x[j];
          x[j]:=x[j+1];
          x[j+1]:=doubleTmp;
        end;
      end;
    end;
  end;

  procedure Tb_assessment.getClassdata(grade : integer; classID : string;var class_ass, grade_ass : assessment);
  var
    sql : string;
    scoreArr : array of double;

    testNum : integer;           //测试人数
    average : double;            //平均值
    dev : double;                //标准差
    greatStuNum : integer;       //优生人数
    inferiorStuNum : integer;    //学困生人数
    teacher : string;            //任课教师

    grateLine : real;            //优秀最低分，含临界
    inferiorLine : double;       //学困的最高分，不含临界

    i : integer;
  begin
    testNum := 0;
    average := 0;
    dev := 0;
    greatStuNum := 0;
    inferiorStuNum := 0;
    teacher := '';
    grateLine := grade_ass.greatLine;
    inferiorLine := grade_ass.inferiorLine;

  //获取这个班该课程的任课教师
    sql := 'SELECT T_'+classID +' as teacher ' +
        ' from  tb_course_'+inttostr(grade)+
        ' where courseName='+#39+courseName+#39;
    ado.SelectInfo(ADOQuery, sql);
    teacher := ADOQuery.FieldByName('teacher').AsString;
  //从数据库中获取这个班该课程所有成绩
    sql := 'select '+examName+' from tb_scores_'+inttostr(grade)+'_'+courseName+
        ' where classID =' +#39+classID+#39;
    ado.SelectInfo(ADOQuery, sql);
    setLength(scoreArr,ADOQuery.RecordCount);

    for I := 0 to ADOQuery.RecordCount-1 do
    begin
      scoreArr[i] := ADOQuery.FieldByName(examName).AsFloat;
      if (scoreArr[i] <> -1) then
      begin
        testNum := testNum+1;
        if (scoreArr[i]>= grateLine) then
          greatStuNum := greatStuNum+1;
        if (scoreArr[i]<inferiorLine) then
          inferiorStuNum := inferiorStuNum+1;
      end;
      ADOQuery.Next;
    end;

    computeDev(scoreArr, average, dev);       //计算平均值和标准差

    if (testNum <> 0) then
    begin
      class_ass.greatRatio := greatStuNum / testNum;
      class_ass.inferiorRatio := inferiorStuNum / testNum;
    end else begin
      class_ass.greatRatio := 0;
      class_ass.inferiorRatio := 0;
    end;
    class_ass.classID := classID;
    class_ass.testNum := testNum;
    class_ass.average := average;
    class_ass.standardDev := dev;
    class_ass.greatStuNum := greatStuNum;
    class_ass.inferiorStuNum := inferiorStuNum;
    class_ass.teacher := teacher;
    class_ass.evaluate := evaluate(grade,class_ass,grade_ass);
  end;

  procedure Tb_assessment.computeDev(var score : array of double; var average, dev : double);
  var
    i : integer;
    sum : double;
    count : integer;  //有效数据个数，非零
  begin
    sum := 0;
    count := 0;

    for I := 0 to Length(score)-1 do
    begin
      if (score[i] <> -1) then      //-1表示没有参加考试
      begin
        sum := sum + score[i];
        count := count + 1;
      end;
    end;

    if count = 0 then
    begin
      average := 0;
      dev := 0;
      exit;
    end else begin
      average := sum / count;

      sum := 0;
      for I := 0 to Length(score)-1 do
      begin
        if (score[i] <> -1) then
        begin
          sum := sum + (score[i]-average)*(score[i]-average);
        end;
      end;

      dev := sqrt(sum / count);
    end;
  end;

  function Tb_assessment.evaluate(grade : integer; class_ass, grade_ass : assessment):string;
  var
    sql : string;
    courseType : string;
    averBase : double;
    averTop : double;
    dev : double;
    greatRatio : double;
//    averVeto : double;
    averNo : double;
    averYes : double;

    averFlag : integer;             //0表示A，1表示B, 2表示C, 3表示一票否决
    dveFlag : integer;              //0表示A，1表示B, 2表示C
    grateFlag : integer;            //0表示A，1表示B, 2表示C

  begin
    //获取参数
    sql := 'SELECT courseType from  tb_course_'+inttostr(grade)+
        ' where courseName='+#39+courseName+#39;
    ado.SelectInfo(ADOQuery, sql);
    courseType := ADOQuery.FieldByName('courseType').AsString;
    sql := 'select averWenBase, averWenTop, averLiBase, averLiTop, averWenNo, averWenYes, averLiNo, averLiYes, dev, greatRatio'+
          ' from tb_parameters  where grade ='+inttostr(grade);
    ado.SelectInfo(ADOQuery, sql);
//    averVeto := ADOQuery.FieldByName('averVeto').Asfloat;

    if Comparestr('文科',courseType)=0 then
    begin
      averBase := ADOQuery.FieldByName('averWenBase').Asfloat;
      averTop := ADOQuery.FieldByName('averWenTop').Asfloat;
      averNo := ADOQuery.FieldByName('averWenNo').Asfloat;
      averYes := ADOQuery.FieldByName('averWenYes').Asfloat;
    end else if Comparestr('理科',courseType)=0 then
    begin
      averBase := ADOQuery.FieldByName('averLiBase').Asfloat;
      averTop := ADOQuery.FieldByName('averLiTop').Asfloat;
      averNo := ADOQuery.FieldByName('averLiNo').Asfloat;
      averYes := ADOQuery.FieldByName('averLiYes').Asfloat;
    end;
    dev := ADOQuery.FieldByName('dev').Asfloat;
    greatRatio := ADOQuery.FieldByName('greatRatio').Asfloat;


    averFlag := evaluateAver(averBase,averTop, averNo, averYes, class_ass.average, grade_ass.average);
    dveFlag := evaluateDev(dev, class_ass.standardDev,grade_ass.standardDev);
    grateFlag := evaluateGrateRatio(greatRatio, class_ass.greatRatio, grade_ass.greatRatio);

    if averFlag = 3 then          //一票否决
    begin
      result := 'C';
    end else if averFlag = -1 then
    begin
      result := 'A';
    end else begin
      if (averFlag = 0) and (dveFlag = 0) and (grateFlag = 0) then
      begin
        result := 'A';
      end else if (averFlag = 2) and (dveFlag = 2) and (grateFlag = 2) then
      begin
        result := 'C';
      end else begin
        result := 'B';
      end;
    end;
  end;

  function Tb_assessment.evaluateAver(base, top, averNo, averYes: double; class_aver, grade_aver: Double):integer;
  begin
    if class_aver < grade_aver + averNo then         //一票否决,averNo是负值
    begin
      result := 3;
    end else if class_aver >=  grade_aver + averYes then     //一票肯定
    begin
      result := -1;
    end else begin
      if class_aver >= (grade_aver + top) then
      begin
        result := 0;
      end else if class_aver >= (grade_aver + base) then   //base是负数
      begin
        result := 1;
      end else begin
        result := 2;
      end;
    end;
  end;

  function Tb_assessment.evaluateDev(dev, class_dev, grade_dev: Double):integer;
  begin
    if class_dev <= (grade_dev - dev) then
    begin
      result := 0;
    end else if (class_dev > (grade_dev + dev)) then
    begin
      result := 2;
    end else begin
      result := 1;
    end;
  end;

  function Tb_assessment.evaluateGrateRatio(greatRatio, class_grateRatio, grade_grateRatio: Double):integer;
  begin
    if class_grateRatio >= (grade_grateRatio + greatRatio*0.01) then
    begin
      result := 0;
    end else if class_grateRatio < (grade_grateRatio - greatRatio*0.01) then
    begin
      result := 2;
    end else begin
      result := 1;
    end;
  end;

  procedure Tb_assessment.initAss(var ass : assessment);
  begin
    ass.classID := '';
    ass.testNum := 0;
    ass.average := 0;
    ass.standardDev := 0;
    ass.greatStuNum := 0;
    ass.greatRatio := 0;
    ass.inferiorStuNum := 0;
    ass.inferiorRatio := 0;
    ass.evaluate := '';
    ass.teacher := '';
    ass.greatLine := 0;
    ass.inferiorLine := 0;
  end;

  procedure Tb_assessment.insertDataToTable(grade : integer; var ass : assessment);
  var
    sql : string;
    trans : CTransform;
  begin

    trans := CTransform.Create;
    ass.classID := trans.classIDToText(ass.classID);

    sql := 'INSERT INTO tb_report_' + examName + '_' + courseName+' Values(' +
          #39 + ass.classID           + #39 + ',' +
          inttostr(ass.testNum )            + ',' +
          Format('%.2f',[ass.average])      + ',' +
          Format('%.2f',[ass.standardDev])  + ',' +
          inttostr(ass.greatStuNum)         + ',' +
          Format('%.2f',[ass.greatRatio*100])   + ',' +
          inttostr(ass.inferiorStuNum)      + ',' +
          Format('%.2f',[ass.inferiorRatio])+ ',' +
          #39 + ass.evaluate          + #39 + ',' +
          #39 + ass.teacher           + #39 + ',' +
          Format('%.1f',[ass.greatLine])    + ',' +
          Format('%.1f',[ass.inferiorLine]) + ')';
    ado.ExecSqlStr(sql);
  end;

end.
