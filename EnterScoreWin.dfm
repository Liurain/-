object Form2: TForm2
  Left = 0
  Top = 0
  Caption = #33606#24030#23454#39564#23567#23398#25104#32489#31649#29702#31995#32479'--'#25104#32489#24405#20837
  ClientHeight = 565
  ClientWidth = 878
  Color = clBtnHighlight
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object control_Panel: TPanel
    Left = 0
    Top = 0
    Width = 249
    Height = 545
    TabOrder = 0
    object Label2: TLabel
      Left = 22
      Top = 159
      Width = 93
      Height = 23
      Caption = #23398'      '#21495#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 22
      Top = 47
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
    object Label1: TLabel
      Left = 21
      Top = 103
      Width = 95
      Height = 23
      Caption = #24405#20837#39034#24207#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 22
      Top = 188
      Width = 93
      Height = 23
      Caption = #22995'      '#21517#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label8: TLabel
      Left = 22
      Top = 18
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
      Left = 22
      Top = 74
      Width = 93
      Height = 23
      Caption = #31185'      '#30446#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label6: TLabel
      Left = 21
      Top = 219
      Width = 93
      Height = 23
      Caption = #25104'      '#32489#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 21
      Top = 129
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
    object grade_ComboBox: TComboBox
      Left = 112
      Top = 20
      Width = 121
      Height = 21
      Style = csDropDownList
      TabOrder = 0
      OnChange = grade_ComboBoxChange
    end
    object classID_ComboBox: TComboBox
      Left = 112
      Top = 47
      Width = 121
      Height = 21
      Style = csDropDownList
      TabOrder = 1
      OnChange = courseAndCalss_ComboBoxChange
    end
    object orderByName_ComboBox: TComboBox
      Left = 112
      Top = 103
      Width = 121
      Height = 21
      Style = csDropDownList
      ItemIndex = 1
      TabOrder = 2
      Text = #23398#21495
      OnChange = orderByName_ComboBoxChange
      Items.Strings = (
        #22995#21517
        #23398#21495)
    end
    object course_ComboBox: TComboBox
      Left = 112
      Top = 74
      Width = 121
      Height = 21
      Style = csDropDownList
      TabOrder = 3
      OnChange = courseAndCalss_ComboBoxChange
    end
    object stuID_Edit: TEdit
      Left = 112
      Top = 159
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 4
    end
    object stuName_Edit: TEdit
      Left = 112
      Top = 186
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 5
    end
    object submit_btn: TButton
      Left = 22
      Top = 250
      Width = 209
      Height = 39
      Caption = #25552#20132
      TabOrder = 6
      OnClick = submit_btnClick
    end
    object score_Edit: TEdit
      Left = 112
      Top = 217
      Width = 121
      Height = 21
      TabOrder = 7
      OnKeyDown = score_EditKeyDown
    end
    object examName_ComboBox: TComboBox
      Left = 112
      Top = 132
      Width = 121
      Height = 21
      Style = csDropDownList
      TabOrder = 8
      OnChange = examName_ComboBoxChange
    end
  end
  object DBGrid1: TDBGrid
    Left = 255
    Top = 8
    Width = 615
    Height = 537
    DataSource = DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnCellClick = DBGrid1CellClick
  end
  object DataSource: TDataSource
    DataSet = ADOQuery
    Left = 664
    Top = 128
  end
  object ADOQuery: TADOQuery
    Parameters = <>
    Left = 664
    Top = 80
  end
end
