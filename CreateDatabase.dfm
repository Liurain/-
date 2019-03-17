object Form4: TForm4
  Left = 0
  Top = 0
  Caption = 'Form4'
  ClientHeight = 299
  ClientWidth = 635
  Color = clWhite
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    Left = 414
    Top = 68
    Width = 188
    Height = 16
    Caption = #20363#22914':2017~2018'#23398#24180#36755#20837'2017'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clSilver
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label21: TLabel
    Left = 175
    Top = 62
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
  object Label1: TLabel
    Left = 175
    Top = 110
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
    Left = 256
    Top = 67
    Width = 145
    Height = 21
    TabOrder = 0
  end
  object inout_submit_btn: TButton
    Left = 175
    Top = 160
    Width = 226
    Height = 49
    Caption = #30830'        '#23450
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -24
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnClick = inout_submit_btnClick
  end
  object term_ComboBox: TComboBox
    Left = 256
    Top = 115
    Width = 145
    Height = 21
    Style = csDropDownList
    TabOrder = 2
    Items.Strings = (
      #19978
      #19979)
  end
end
