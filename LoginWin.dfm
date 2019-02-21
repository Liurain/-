object Form3: TForm3
  Left = 0
  Top = 0
  Caption = #33606#24030#24066#23454#39564#23567#23398#25104#32489#31649#29702#31995#32479
  ClientHeight = 319
  ClientWidth = 535
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel: TPanel
    Left = 0
    Top = 0
    Width = 537
    Height = 321
    Color = clBtnHighlight
    ParentBackground = False
    TabOrder = 0
    object Label23: TLabel
      Left = 89
      Top = 89
      Width = 57
      Height = 23
      Caption = #25968#25454#24211
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label24: TLabel
      Left = 90
      Top = 129
      Width = 56
      Height = 23
      Caption = #36134'   '#21495
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label25: TLabel
      Left = 90
      Top = 172
      Width = 56
      Height = 23
      Caption = #23494'   '#30721
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object login_btn: TButton
      Left = 256
      Top = 217
      Width = 161
      Height = 41
      Caption = #30331#38470
      TabOrder = 0
      OnClick = login_btnClick
    end
    object user_Edit: TEdit
      Left = 152
      Top = 134
      Width = 265
      Height = 21
      TabOrder = 1
      Text = 'admin'
    end
    object psw_Edit: TEdit
      Left = 152
      Top = 175
      Width = 265
      Height = 21
      PasswordChar = '*'
      TabOrder = 2
      Text = '123'
    end
    object dataPath_ComboBox: TComboBox
      Left = 152
      Top = 89
      Width = 265
      Height = 21
      Style = csDropDownList
      TabOrder = 3
    end
    object creatDatabase_btn: TButton
      Left = 89
      Top = 217
      Width = 161
      Height = 41
      Caption = #21019#24314#26032#25968#25454#24211
      TabOrder = 4
      OnClick = creatDatabase_btnClick
    end
  end
end
