unit GlobalData;

interface

var
  {****************整个程序的全局变量*******************}
  averWenBase : real;         //文科学科评定为B级别，低于年级平均分的最大分数
  averWenTop :real;           //文科学科评定为B级别，高于年级平均分的最大分数
  averLiBase : real;          //理科学科评定为B级别，低于年级平均分的最大分数
  averLiTop : real;           //理科学科评定为B级别，高于年级平均分的最大分数
  dev : real;                 //标准差不大于该参数评为A、小于评委C
  greatRatio : real;          //优生率高于该百分比评为A、小于该百分比评委C
  year : integer;             //开学年份，例如2017~2018学年为2017

  power : array[1..9] of boolean;       //当前身份是否能够访问成绩的权限


implementation



end.
