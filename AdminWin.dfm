object Form1: TForm1
  Left = 0
  Top = 0
  Caption = #33606#24030#23454#39564#23567#23398#25104#32489#31649#29702#31995#32479'--'#31649#29702#21592
  ClientHeight = 587
  ClientWidth = 870
  Color = clBtnHighlight
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object Label51: TLabel
    Left = 25
    Top = 173
    Width = 132
    Height = 23
    Caption = #24179' '#22343' '#20998' '#21542#20915#65306
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label52: TLabel
    Left = 16
    Top = 277
    Width = 132
    Height = 23
    Caption = #24179' '#22343' '#20998' '#21542#20915#65306
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 2
    Width = 313
    Height = 577
    ActivePage = databasePage
    TabOrder = 0
    object classPage: TTabSheet
      Caption = #29677#32423#31649#29702
      object PanelControlClass: TPanel
        Left = 15
        Top = 65
        Width = 274
        Height = 145
        BevelOuter = bvNone
        TabOrder = 0
        object Label2: TLabel
          Left = 15
          Top = 10
          Width = 75
          Height = 23
          Caption = #24180'   '#32423#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label3: TLabel
          Left = 16
          Top = 39
          Width = 75
          Height = 23
          Caption = #29677'   '#21495#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label4: TLabel
          Left = 16
          Top = 68
          Width = 75
          Height = 23
          Caption = #20154'   '#25968#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 16
          Top = 97
          Width = 76
          Height = 23
          Caption = #29677#20027#20219#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object class_grade_ComboBox: TComboBox
          Left = 96
          Top = 10
          Width = 121
          Height = 21
          Style = csDropDownList
          TabOrder = 0
          OnChange = class_grade_ComboBoxChange
          Items.Strings = (
            #19968#24180#32423
            #20108#24180#32423
            #19977#24180#32423
            #22235#24180#32423
            #20116#24180#32423
            #20845#24180#32423
            #19971#24180#32423
            #20843#24180#32423
            #20061#24180#32423)
        end
        object class_classID_ComboBox: TComboBox
          Left = 96
          Top = 37
          Width = 121
          Height = 21
          TabOrder = 1
        end
        object class_studentsNum_Edit: TEdit
          Left = 97
          Top = 68
          Width = 121
          Height = 21
          Enabled = False
          TabOrder = 2
        end
        object class_teacher_Edit: TEdit
          Left = 97
          Top = 97
          Width = 121
          Height = 21
          TabOrder = 3
        end
      end
      object class_operate_RG: TRadioGroup
        Left = 15
        Top = 3
        Width = 234
        Height = 43
        Caption = #25805#20316#36873#25321
        Columns = 3
        Items.Strings = (
          #22686#21152
          #20462#25913
          #21024#38500)
        TabOrder = 1
        OnClick = class_operate_RGClick
      end
      object class_submit_btn: TButton
        Left = 31
        Top = 216
        Width = 209
        Height = 39
        Caption = #25552#20132
        TabOrder = 2
        OnClick = class_submit_btnClick
      end
    end
    object coursePage: TTabSheet
      Caption = #35838#31243#31649#29702
      ImageIndex = 1
      object PanelControlCourse: TPanel
        Left = 11
        Top = 52
        Width = 235
        Height = 190
        BevelOuter = bvNone
        TabOrder = 0
        object Label6: TLabel
          Left = 8
          Top = 41
          Width = 95
          Height = 23
          Caption = #35838#31243#21517#31216#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label8: TLabel
          Left = 8
          Top = 125
          Width = 95
          Height = 23
          Caption = #20219#35838#25945#24072#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label1: TLabel
          Left = 8
          Top = 10
          Width = 93
          Height = 23
          Caption = #24180'      '#32423#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label9: TLabel
          Left = 10
          Top = 97
          Width = 93
          Height = 23
          Caption = #29677'      '#21495#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label22: TLabel
          Left = 10
          Top = 70
          Width = 95
          Height = 23
          Caption = #35838#31243#31867#21035#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object course_courseName_Edit: TEdit
          Left = 97
          Top = 39
          Width = 128
          Height = 21
          TabOrder = 0
        end
        object course_teacher_Edit: TEdit
          Left = 97
          Top = 125
          Width = 128
          Height = 21
          TabOrder = 1
        end
        object course_grade_ComboBox: TComboBox
          Left = 97
          Top = 10
          Width = 128
          Height = 21
          Style = csDropDownList
          TabOrder = 2
          OnChange = grade_ComboBoxChange
          Items.Strings = (
            #19968#24180#32423
            #20108#24180#32423
            #19977#24180#32423
            #22235#24180#32423
            #20116#24180#32423
            #20845#24180#32423
            #19971#24180#32423
            #20843#24180#32423
            #20061#24180#32423)
        end
        object course_classID_ComboBox: TComboBox
          Left = 97
          Top = 95
          Width = 128
          Height = 21
          Style = csDropDownList
          TabOrder = 3
          OnChange = course_classID_ComboBoxChange
        end
        object course_courseType_ComboBox: TComboBox
          Left = 97
          Top = 68
          Width = 128
          Height = 21
          Style = csDropDownList
          TabOrder = 4
          Items.Strings = (
            #29702#31185
            #25991#31185)
        end
      end
      object course_submit_btn: TButton
        Left = 27
        Top = 248
        Width = 209
        Height = 39
        Caption = #25552#20132
        TabOrder = 1
        OnClick = course_submit_btnClick
      end
      object course_operate_RG: TRadioGroup
        Left = 13
        Top = 3
        Width = 233
        Height = 43
        Caption = #25805#20316#36873#25321
        Columns = 3
        Items.Strings = (
          #22686#21152
          #20462#25913
          #21024#38500)
        TabOrder = 2
        OnClick = course_operate_RGClick
      end
    end
    object studentPage: TTabSheet
      Caption = #23398#29983#31649#29702
      ImageIndex = 2
      object PanelControlStudents: TPanel
        Left = 3
        Top = 49
        Width = 246
        Height = 256
        BevelOuter = bvNone
        TabOrder = 0
        object Label12: TLabel
          Left = 24
          Top = 92
          Width = 75
          Height = 23
          Caption = #23398'   '#21495#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label13: TLabel
          Left = 24
          Top = 121
          Width = 75
          Height = 23
          Caption = #22995'   '#21517#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label14: TLabel
          Left = 24
          Top = 150
          Width = 75
          Height = 23
          Caption = #24615'   '#21035#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label15: TLabel
          Left = 24
          Top = 179
          Width = 75
          Height = 23
          Caption = #24180'   '#40836#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label16: TLabel
          Left = 24
          Top = 10
          Width = 75
          Height = 23
          Caption = #24180'   '#32423#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label17: TLabel
          Left = 24
          Top = 39
          Width = 75
          Height = 23
          Caption = #29677'   '#21495#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label45: TLabel
          Left = 24
          Top = 65
          Width = 75
          Height = 23
          Caption = #20154'   '#25968#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object stu_sex_ComboBox: TComboBox
          Left = 105
          Top = 150
          Width = 128
          Height = 21
          Style = csDropDownList
          ItemIndex = 0
          TabOrder = 0
          Text = #30007
          Items.Strings = (
            #30007
            #22899)
        end
        object stu_stuID_Edit: TEdit
          Left = 105
          Top = 92
          Width = 128
          Height = 21
          TabOrder = 1
        end
        object stu_name_Edit: TEdit
          Left = 105
          Top = 121
          Width = 128
          Height = 21
          TabOrder = 2
        end
        object stu_grade_ComboBox: TComboBox
          Left = 105
          Top = 12
          Width = 128
          Height = 21
          Style = csDropDownList
          TabOrder = 3
          OnChange = grade_ComboBoxChange
          Items.Strings = (
            #19968#24180#32423
            #20108#24180#32423
            #19977#24180#32423
            #22235#24180#32423
            #20116#24180#32423
            #20845#24180#32423
            #19971#24180#32423
            #20843#24180#32423
            #20061#24180#32423)
        end
        object stu_classID_ComboBox: TComboBox
          Left = 105
          Top = 39
          Width = 128
          Height = 21
          Style = csDropDownList
          TabOrder = 4
          OnChange = stu_classID_ComboBoxChange
        end
        object stu_age_ComboBox: TComboBox
          Left = 105
          Top = 179
          Width = 128
          Height = 21
          TabOrder = 5
          Items.Strings = (
            '6'
            '7'
            '8'
            '9'
            '10'
            '11'
            '12'
            '13'
            '14')
        end
        object stu_stuNum_Edit: TEdit
          Left = 105
          Top = 65
          Width = 128
          Height = 21
          Enabled = False
          TabOrder = 6
        end
      end
      object stu_submit_btn: TButton
        Left = 27
        Top = 327
        Width = 209
        Height = 39
        Caption = #25552#20132
        TabOrder = 1
        OnClick = stu_submit_btnClick
      end
      object stu_operate_RG: TRadioGroup
        Left = 13
        Top = 0
        Width = 236
        Height = 43
        Caption = #25805#20316#36873#25321
        Columns = 3
        Items.Strings = (
          #22686#21152
          #20462#25913
          #21024#38500)
        TabOrder = 2
        OnClick = stu_operate_RGClick
      end
    end
    object parPage: TTabSheet
      Caption = #21442#25968#37197#32622
      ImageIndex = 3
      object Panel1: TPanel
        Left = 3
        Top = 0
        Width = 299
        Height = 497
        BevelOuter = bvNone
        TabOrder = 0
        object Label7: TLabel
          Left = 9
          Top = 99
          Width = 133
          Height = 23
          Caption = #25991#31185#22343#20998#21442#25968#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label10: TLabel
          Left = 9
          Top = 128
          Width = 133
          Height = 23
          Caption = #29702#31185#22343#20998#21442#25968#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label11: TLabel
          Left = 9
          Top = 266
          Width = 132
          Height = 23
          Caption = #26631' '#20934' '#24046' '#21442#25968#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label33: TLabel
          Left = 9
          Top = 295
          Width = 133
          Height = 23
          Caption = #20248' '#29983' '#29575' (%) '#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label36: TLabel
          Left = 34
          Top = 12
          Width = 230
          Height = 35
          Caption = 'B'#32423#35780#23450#21442#25968#35774#23450
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -29
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object Label39: TLabel
          Left = 199
          Top = 101
          Width = 17
          Height = 23
          Caption = #8212
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label42: TLabel
          Left = 199
          Top = 126
          Width = 17
          Height = 23
          Caption = #8212
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label41: TLabel
          Left = 9
          Top = 70
          Width = 133
          Height = 23
          Caption = #21442#25968#23545#24212#24180#32423#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label43: TLabel
          Left = 9
          Top = 320
          Width = 134
          Height = 23
          Caption = #20248#29983#21344#27604'(%) '#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label44: TLabel
          Left = 9
          Top = 349
          Width = 128
          Height = 23
          Caption = #23398#22256#21344#27604'(%) : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label49: TLabel
          Left = 9
          Top = 157
          Width = 133
          Height = 23
          Caption = #25991#31185#22343#20998#21542#20915#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label37: TLabel
          Left = 9
          Top = 240
          Width = 133
          Height = 23
          Caption = #29702#31185#22343#20998#32943#23450#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label53: TLabel
          Left = 7
          Top = 184
          Width = 133
          Height = 23
          Caption = #25991#31185#22343#20998#32943#23450#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label54: TLabel
          Left = 9
          Top = 211
          Width = 133
          Height = 23
          Caption = #29702#31185#22343#20998#21542#20915#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object par_averLi_base_Edit: TEdit
          Left = 137
          Top = 130
          Width = 60
          Height = 21
          TabOrder = 0
        end
        object par_dev_Edit: TEdit
          Left = 137
          Top = 266
          Width = 140
          Height = 21
          TabOrder = 1
        end
        object par_greatRatio_Edit: TEdit
          Left = 137
          Top = 297
          Width = 140
          Height = 21
          TabOrder = 2
        end
        object par_averWen_base_Edit: TEdit
          Left = 137
          Top = 103
          Width = 60
          Height = 21
          TabOrder = 3
        end
        object par_averWen_top_Edit: TEdit
          Left = 217
          Top = 103
          Width = 60
          Height = 21
          TabOrder = 4
        end
        object par_averLi_top_Edit: TEdit
          Left = 217
          Top = 130
          Width = 60
          Height = 21
          TabOrder = 5
        end
        object par_grade_ComboBox: TComboBox
          Left = 137
          Top = 72
          Width = 140
          Height = 21
          Style = csDropDownList
          TabOrder = 6
          Items.Strings = (
            #19968#24180#32423
            #20108#24180#32423
            #19977#24180#32423
            #22235#24180#32423
            #20116#24180#32423
            #20845#24180#32423
            #19971#24180#32423
            #20843#24180#32423
            #20061#24180#32423)
        end
        object par_greatWeight_Edit: TEdit
          Left = 137
          Top = 324
          Width = 140
          Height = 21
          TabOrder = 7
        end
        object par_inferiorWeight_Edit: TEdit
          Left = 137
          Top = 351
          Width = 140
          Height = 21
          TabOrder = 8
        end
        object par_averWenNo_Edit: TEdit
          Left = 137
          Top = 157
          Width = 140
          Height = 21
          TabOrder = 9
        end
        object par_averLiYes_Edit: TEdit
          Left = 137
          Top = 238
          Width = 140
          Height = 21
          TabOrder = 10
        end
        object par_averWenYes_Edit: TEdit
          Left = 137
          Top = 184
          Width = 140
          Height = 21
          TabOrder = 11
        end
        object par_averLiNo_Edit: TEdit
          Left = 137
          Top = 211
          Width = 140
          Height = 21
          TabOrder = 12
        end
        object par_submit_btn: TButton
          Left = 34
          Top = 412
          Width = 209
          Height = 39
          Caption = #25552#20132
          TabOrder = 13
          OnClick = par_submit_btnClick
        end
      end
    end
    object outputPage: TTabSheet
      Caption = #29983#25104#25253#21578
      ImageIndex = 4
      object Panel2: TPanel
        Left = 3
        Top = 1
        Width = 302
        Height = 481
        BevelOuter = bvNone
        TabOrder = 0
        object report_submit_btn: TButton
          Left = 48
          Top = 384
          Width = 193
          Height = 49
          Caption = #30830#23450
          TabOrder = 0
          OnClick = report_submit_btnClick
        end
        object fileOutputPanel: TPanel
          Left = -3
          Top = 0
          Width = 292
          Height = 324
          BevelOuter = bvNone
          TabOrder = 1
          object Label23: TLabel
            Left = 13
            Top = 15
            Width = 95
            Height = 23
            Caption = #25253#21578#31867#22411#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label24: TLabel
            Left = 13
            Top = 191
            Width = 93
            Height = 23
            Caption = #24180'      '#32423#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label25: TLabel
            Left = 13
            Top = 222
            Width = 93
            Height = 23
            Caption = #29677'      '#21495#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label26: TLabel
            Left = 13
            Top = 91
            Width = 95
            Height = 23
            Caption = #31185#30446#31867#22411#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label38: TLabel
            Left = 13
            Top = 164
            Width = 95
            Height = 23
            Caption = #32771#35797#21517#31216#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label46: TLabel
            Left = 13
            Top = 46
            Width = 95
            Height = 23
            Caption = #25991#20214#31867#22411#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label47: TLabel
            Left = 13
            Top = 249
            Width = 95
            Height = 23
            Caption = #35838#31243#21517#31216#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label48: TLabel
            Left = 13
            Top = 120
            Width = 95
            Height = 23
            Caption = #36755#20986#33539#22260#65306
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -19
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object report_type_ComboBox: TComboBox
            Left = 104
            Top = 19
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 0
            OnChange = report_type_ComboBoxChange
            Items.Strings = (
              #30331#20998#34920
              #25104#32489#34920
              #26399#26411#35780#20272
              #32771#35797#35780#20272)
          end
          object report_grade_ComboBox: TComboBox
            Left = 104
            Top = 195
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 1
            OnChange = report_grade_ComboBoxChange
            Items.Strings = (
              #19968#24180#32423
              #20108#24180#32423
              #19977#24180#32423
              #22235#24180#32423
              #20116#24180#32423
              #20845#24180#32423
              #19971#24180#32423
              #20843#24180#32423
              #20061#24180#32423)
          end
          object report_classID_ComboBox: TComboBox
            Left = 104
            Top = 222
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 2
          end
          object report_courseType_ComboBox: TComboBox
            Left = 104
            Top = 93
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 3
            OnChange = report_courseType_ComboBoxChange
            Items.Strings = (
              #20840#31185
              #21333#31185)
          end
          object report_exam_ComboBox: TComboBox
            Left = 104
            Top = 164
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 4
            OnChange = report_exam_ComboBoxChange
          end
          object visual_CheckBox: TCheckBox
            Left = 40
            Top = 292
            Width = 97
            Height = 17
            Caption = #29983#25104#25253#21578#21518#25171#24320
            TabOrder = 5
          end
          object fileprint_CheckBox: TCheckBox
            Left = 168
            Top = 292
            Width = 97
            Height = 17
            Caption = #25171#21360#25991#20214
            TabOrder = 6
          end
          object report_Filetype_ComboBox: TComboBox
            Left = 104
            Top = 46
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 7
            OnChange = report_Filetype_ComboBoxChange
            Items.Strings = (
              'Word'
              'Excel')
          end
          object report_courseName_ComboBox: TComboBox
            Left = 104
            Top = 249
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 8
          end
          object report_outPutLimits_ComboBox: TComboBox
            Left = 104
            Top = 120
            Width = 137
            Height = 21
            Style = csDropDownList
            TabOrder = 9
            OnChange = report_outPutLimits_ComboBoxChange
            Items.Strings = (
              #20840#26657
              #26576#19968#24180#32423
              #26576#19968#29677#32423)
          end
        end
      end
    end
    object AdminPage: TTabSheet
      Caption = 'AdminPage'
      ImageIndex = 5
      object userPanel: TPanel
        Left = 15
        Top = 0
        Width = 306
        Height = 534
        BevelOuter = bvNone
        TabOrder = 0
        object Label27: TLabel
          Left = 20
          Top = 30
          Width = 64
          Height = 19
          Caption = #29992' '#25143' '#21517':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label28: TLabel
          Left = 20
          Top = 57
          Width = 64
          Height = 19
          Caption = #26087' '#23494' '#30721':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label29: TLabel
          Left = 20
          Top = 87
          Width = 64
          Height = 19
          Caption = #26032' '#23494' '#30721':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label30: TLabel
          Left = 20
          Top = 113
          Width = 70
          Height = 19
          Caption = #23494#30721#30830#35748':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object admin_newPsw_Edit: TEdit
          Left = 96
          Top = 86
          Width = 169
          Height = 21
          HelpType = htKeyword
          PasswordChar = '*'
          TabOrder = 0
        end
        object admin_checkPsw_Edit: TEdit
          Left = 96
          Top = 113
          Width = 169
          Height = 21
          HelpType = htKeyword
          HelpKeyword = '*'
          PasswordChar = '*'
          TabOrder = 1
        end
        object admin_oldPsw_Edit: TEdit
          Left = 96
          Top = 59
          Width = 169
          Height = 21
          TabOrder = 2
        end
        object admin_userName_Edit: TEdit
          Left = 96
          Top = 31
          Width = 169
          Height = 21
          Enabled = False
          TabOrder = 3
          Text = 'admin'
        end
        object admin_submit_btn: TButton
          Left = 48
          Top = 160
          Width = 201
          Height = 57
          Caption = #20462#25913
          TabOrder = 4
          OnClick = admin_submit_btnClick
        end
      end
    end
    object examPage: TTabSheet
      Caption = 'examPage'
      ImageIndex = 6
      object Panel4: TPanel
        Left = 0
        Top = -5
        Width = 305
        Height = 529
        BevelOuter = bvNone
        TabOrder = 0
        object Label34: TLabel
          Left = 28
          Top = 86
          Width = 70
          Height = 19
          Caption = #32771#35797#21517#31216':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label35: TLabel
          Left = 28
          Top = 113
          Width = 70
          Height = 19
          Caption = #32771#35797#26102#38388':'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object exam_operate_RG: TRadioGroup
          Left = 29
          Top = 16
          Width = 236
          Height = 43
          Caption = #25805#20316#36873#25321
          Columns = 3
          Items.Strings = (
            #22686#21152
            #20462#25913
            #21024#38500)
          TabOrder = 0
          OnClick = exam_operate_RGClick
        end
        object examTime_Edit: TEdit
          Left = 104
          Top = 115
          Width = 169
          Height = 21
          Enabled = False
          TabOrder = 1
        end
        object examName_Edit: TEdit
          Left = 104
          Top = 88
          Width = 169
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clSilver
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          Text = #35831#22635#20889#23613#37327#31616#30701#30340#27721#23383
          OnClick = examName_EditClick
        end
        object Panel5: TPanel
          Left = 12
          Top = 138
          Width = 278
          Height = 185
          BevelOuter = bvNone
          TabOrder = 3
          object examGradePanel: TPanel
            Left = 16
            Top = 79
            Width = 257
            Height = 97
            BevelOuter = bvNone
            TabOrder = 0
            object examGrade8_CheckBox: TCheckBox
              Left = 97
              Top = 71
              Width = 57
              Height = 17
              Caption = #20843#24180#32423
              TabOrder = 0
            end
            object examGrade7_CheckBox: TCheckBox
              Left = 16
              Top = 71
              Width = 58
              Height = 17
              Caption = #19971#24180#32423
              TabOrder = 1
            end
            object examGrade9_CheckBox: TCheckBox
              Left = 176
              Top = 72
              Width = 57
              Height = 17
              Caption = #20061#24180#32423
              TabOrder = 2
            end
            object examGrade2_CheckBox: TCheckBox
              Left = 97
              Top = 8
              Width = 57
              Height = 17
              Caption = #20108#24180#32423
              TabOrder = 3
            end
            object examGrade1_CheckBox: TCheckBox
              Left = 16
              Top = 7
              Width = 58
              Height = 17
              Caption = #19968#24180#32423
              TabOrder = 4
            end
            object examGrade3_CheckBox: TCheckBox
              Left = 176
              Top = 8
              Width = 57
              Height = 17
              Caption = #19977#24180#32423
              TabOrder = 5
            end
            object examGrade5_CheckBox: TCheckBox
              Left = 97
              Top = 40
              Width = 57
              Height = 17
              Caption = #20116#24180#32423
              TabOrder = 6
            end
            object examGrade4_CheckBox: TCheckBox
              Left = 16
              Top = 39
              Width = 58
              Height = 17
              Caption = #22235#24180#32423
              TabOrder = 7
            end
            object examGrade6_CheckBox: TCheckBox
              Left = 176
              Top = 40
              Width = 57
              Height = 17
              Caption = #20845#24180#32423
              TabOrder = 8
            end
          end
          object exam_grade_RG: TRadioGroup
            Left = 24
            Top = 8
            Width = 249
            Height = 65
            Caption = #21442#32771#24180#32423
            Columns = 3
            Items.Strings = (
              #25152#26377#24180#32423
              '1~3'#24180#32423
              '4~6'#24180#32423
              '7~9'#24180#32423
              #33258#23450#20041)
            TabOrder = 1
            OnClick = exam_grade_RGClick
          end
        end
        object exam_submit_btn: TButton
          Left = 44
          Top = 343
          Width = 241
          Height = 49
          Caption = #30830#35748
          TabOrder = 4
          OnClick = exam_submit_btnClick
        end
      end
    end
    object usersPage: TTabSheet
      Caption = 'usersPage'
      ImageIndex = 7
      object Label31: TLabel
        Left = 20
        Top = 78
        Width = 64
        Height = 19
        Caption = #29992' '#25143' '#21517':'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object Label32: TLabel
        Left = 20
        Top = 105
        Width = 63
        Height = 19
        Caption = #23494'     '#30721':'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object user_operate_RG: TRadioGroup
        Left = 21
        Top = 8
        Width = 236
        Height = 43
        Caption = #25805#20316#36873#25321
        Columns = 3
        Items.Strings = (
          #22686#21152
          #20462#25913
          #21024#38500)
        TabOrder = 0
        OnClick = user_operate_RGClick
      end
      object userPsw_Edit: TEdit
        Left = 96
        Top = 107
        Width = 169
        Height = 21
        TabOrder = 1
      end
      object userName_Edit: TEdit
        Left = 96
        Top = 79
        Width = 169
        Height = 21
        Enabled = False
        TabOrder = 2
      end
      object powerPanel: TPanel
        Left = 3
        Top = 134
        Width = 278
        Height = 185
        BevelOuter = bvNone
        TabOrder = 3
        object userPowerPanel: TPanel
          Left = 17
          Top = 74
          Width = 257
          Height = 97
          BevelOuter = bvNone
          TabOrder = 0
          object userPower8_CheckBox: TCheckBox
            Left = 92
            Top = 72
            Width = 57
            Height = 17
            Caption = #20843#24180#32423
            TabOrder = 0
          end
          object userPower7_CheckBox: TCheckBox
            Left = 11
            Top = 71
            Width = 58
            Height = 17
            Caption = #19971#24180#32423
            TabOrder = 1
          end
          object userPower9_CheckBox: TCheckBox
            Left = 171
            Top = 71
            Width = 57
            Height = 17
            Caption = #20061#24180#32423
            TabOrder = 2
          end
          object userPower2_CheckBox: TCheckBox
            Left = 92
            Top = 8
            Width = 57
            Height = 17
            Caption = #20108#24180#32423
            TabOrder = 3
          end
          object userPower1_CheckBox: TCheckBox
            Left = 11
            Top = 7
            Width = 58
            Height = 17
            Caption = #19968#24180#32423
            TabOrder = 4
          end
          object userPower3_CheckBox: TCheckBox
            Left = 171
            Top = 8
            Width = 57
            Height = 17
            Caption = #19977#24180#32423
            TabOrder = 5
          end
          object userPower5_CheckBox: TCheckBox
            Left = 92
            Top = 40
            Width = 57
            Height = 17
            Caption = #20116#24180#32423
            TabOrder = 6
          end
          object userPower4_CheckBox: TCheckBox
            Left = 11
            Top = 39
            Width = 58
            Height = 17
            Caption = #22235#24180#32423
            TabOrder = 7
          end
          object userPower6_CheckBox: TCheckBox
            Left = 171
            Top = 40
            Width = 57
            Height = 17
            Caption = #20845#24180#32423
            TabOrder = 8
          end
        end
        object user_power_RG: TRadioGroup
          Left = 17
          Top = 3
          Width = 255
          Height = 65
          Caption = #26435#38480
          Columns = 3
          Items.Strings = (
            #25152#26377#24180#32423
            '1~3'#24180#32423
            '4~6'#24180#32423
            '7~9'#24180#32423
            #33258#23450#20041)
          TabOrder = 1
          OnClick = user_power_RGClick
        end
      end
      object user_submit_btn: TButton
        Left = 40
        Top = 360
        Width = 241
        Height = 49
        Caption = #30830#35748
        TabOrder = 4
        OnClick = user_submit_btnClick
      end
    end
    object inputPage: TTabSheet
      Caption = 'inputPage'
      ImageIndex = 8
      object Panel6: TPanel
        Left = 10
        Top = 0
        Width = 292
        Height = 193
        BevelOuter = bvNone
        TabOrder = 0
        object Label18: TLabel
          Left = 13
          Top = 15
          Width = 95
          Height = 23
          Caption = #25968#25454#31867#22411#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label19: TLabel
          Left = 13
          Top = 71
          Width = 93
          Height = 23
          Caption = #25991'      '#20214#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label40: TLabel
          Left = 13
          Top = 42
          Width = 95
          Height = 23
          Caption = #32771#35797#21517#31216#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object input_type_ComboBox: TComboBox
          Left = 104
          Top = 17
          Width = 145
          Height = 21
          Style = csDropDownList
          TabOrder = 0
          OnChange = input_type_ComboBoxChange
          Items.Strings = (
            #35838#31243
            #29677#32423
            #23398#29983
            #25104#32489)
        end
        object input_filePath_Edit: TEdit
          Left = 102
          Top = 76
          Width = 121
          Height = 21
          TabOrder = 1
        end
        object input_getFile_Button: TButton
          Left = 224
          Top = 74
          Width = 25
          Height = 25
          Caption = '...'
          TabOrder = 2
          OnClick = input_getFile_ButtonClick
        end
        object input_exam_ComboBox: TComboBox
          Left = 104
          Top = 44
          Width = 145
          Height = 21
          Style = csDropDownList
          TabOrder = 3
          OnChange = report_type_ComboBoxChange
        end
      end
      object inout_submit_btn: TButton
        Left = 64
        Top = 240
        Width = 169
        Height = 49
        Caption = #30830#23450
        TabOrder = 1
        OnClick = inout_submit_btnClick
      end
    end
    object databasePage: TTabSheet
      Caption = 'databasePage'
      ImageIndex = 9
      object Panel7: TPanel
        Left = 13
        Top = 15
        Width = 292
        Height = 334
        BevelOuter = bvNone
        TabOrder = 0
        object Label21: TLabel
          Left = 0
          Top = 17
          Width = 75
          Height = 23
          Caption = #23398'   '#24180#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label50: TLabel
          Left = 0
          Top = 46
          Width = 75
          Height = 23
          Caption = #23398'   '#26399#65306
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object database_beginYear_Edit: TEdit
          Left = 73
          Top = 17
          Width = 112
          Height = 21
          TabOrder = 0
        end
        object keepBack_CheckBox: TCheckBox
          Left = 191
          Top = 48
          Width = 97
          Height = 17
          Caption = #20445#30041#20197#24448#25968#25454
          TabOrder = 1
          OnClick = keepBack_CheckBoxClick
        end
        object Panel8: TPanel
          Left = 0
          Top = 86
          Width = 289
          Height = 161
          BevelOuter = bvNone
          TabOrder = 2
          object Label20: TLabel
            Left = 0
            Top = 8
            Width = 154
            Height = 25
            Caption = #20801#35768#20445#30041#30340#25968#25454
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -21
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
          end
          object data_course_CheckBox: TCheckBox
            Left = 32
            Top = 56
            Width = 97
            Height = 17
            Caption = #35838#31243#35774#32622
            Enabled = False
            TabOrder = 0
          end
          object data_par_CheckBox: TCheckBox
            Left = 152
            Top = 56
            Width = 97
            Height = 17
            Caption = #21442#25968#37197#32622
            Enabled = False
            TabOrder = 1
          end
          object data_teacher_CheckBox: TCheckBox
            Left = 32
            Top = 79
            Width = 97
            Height = 17
            Caption = #29677#20027#20219#20219#21629
            Enabled = False
            TabOrder = 2
          end
          object data_stu_CheckBox: TCheckBox
            Left = 32
            Top = 102
            Width = 193
            Height = 17
            Caption = #23398#29983#21517#20876'('#36328#23398#24180#26102#19981#21547#27605#19994#29677#65289
            Enabled = False
            TabOrder = 3
          end
          object data_user_CheckBox: TCheckBox
            Left = 152
            Top = 79
            Width = 97
            Height = 17
            Caption = #29992#25143#35774#23450
            Enabled = False
            TabOrder = 4
          end
        end
        object term_ComboBox: TComboBox
          Left = 73
          Top = 46
          Width = 112
          Height = 21
          Style = csDropDownList
          TabOrder = 3
          OnChange = term_ComboBoxChange
          Items.Strings = (
            #19978
            #19979)
        end
      end
      object database_submit_btn: TButton
        Left = 58
        Top = 368
        Width = 169
        Height = 49
        Caption = #30830#23450
        TabOrder = 1
        OnClick = database_submit_btnClick
      end
    end
  end
  object tab_Panel: TPanel
    Left = 640
    Top = 21
    Width = 153
    Height = 187
    BevelOuter = bvNone
    TabOrder = 1
  end
  object promptPanel: TPanel
    Left = 454
    Top = 503
    Width = 408
    Height = 76
    BevelOuter = bvNone
    TabOrder = 2
    object promptLabel: TLabel
      Left = 20
      Top = 17
      Width = 357
      Height = 42
      Caption = #25968#25454#23548#20986#20013#65292#35831#31245#21518'...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -35
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object PageControl2: TPageControl
    Left = 345
    Top = 128
    Width = 289
    Height = 297
    ActivePage = TabSheet
    TabOrder = 3
    object TabSheet: TTabSheet
      Caption = 'TabSheet'
      object DBGrid1: TDBGrid
        Left = 0
        Top = 20
        Width = 113
        Height = 115
        DataSource = DataSource
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid1CellClick
      end
    end
    object TextSheet: TTabSheet
      Caption = 'TextSheet'
      ImageIndex = 1
      object Memo1: TMemo
        Left = 0
        Top = 0
        Width = 193
        Height = 121
        ScrollBars = ssBoth
        TabOrder = 0
      end
    end
  end
  object DataSource: TDataSource
    DataSet = ADOQuery
    Left = 816
    Top = 152
  end
  object ADOQuery: TADOQuery
    Parameters = <>
    Left = 816
    Top = 80
  end
  object MainMenu1: TMainMenu
    Left = 816
    Top = 16
    object N1: TMenuItem
      Caption = #25991#20214
      object dataInput_Menu: TMenuItem
        Caption = #25968#25454#23548#20837
        OnClick = dataInput_MenuClick
      end
      object dataOutput_Menu: TMenuItem
        Caption = #29983#25104#25253#34920
        OnClick = dataOutput_MenuClick
      end
    end
    object N4: TMenuItem
      Caption = #25968#25454#31649#29702
      object class_Menu: TMenuItem
        Caption = #29677#32423#31649#29702
        OnClick = class_MenuClick
      end
      object course_Menu: TMenuItem
        Caption = #35838#31243#31649#29702
        OnClick = course_MenuClick
      end
      object stu_Menu: TMenuItem
        Caption = #23398#29983#31649#29702
        OnClick = stu_MenuClick
      end
      object N2: TMenuItem
        Caption = #25104#32489#31649#29702
        OnClick = N2Click
      end
    end
    object exam_Menu: TMenuItem
      Caption = #32771#21153#31649#29702
      OnClick = exam_MenuClick
    end
    object N10: TMenuItem
      Caption = #29992#25143#31649#29702
      object scoreInput_Menu: TMenuItem
        Caption = #25104#32489#24405#20837#20154#21592
        OnClick = scoreInput_MenuClick
      end
      object adminInfo_Menu: TMenuItem
        Caption = #31649#29702#21592
        OnClick = adminInfo_MenuClick
      end
    end
    object par_Menu: TMenuItem
      Caption = #21442#25968#37197#32622
      OnClick = par_MenuClick
    end
    object database_Menu: TMenuItem
      Caption = #25968#25454#24211#31649#29702
      OnClick = database_MenuClick
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 816
    Top = 288
  end
end
