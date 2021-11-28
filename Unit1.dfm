object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Xred'#34837#34411#28165#29702#24037#20855'@forTheBest'
  ClientHeight = 490
  ClientWidth = 806
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = UnInit
  OnCreate = Init
  PixelsPerInch = 96
  TextHeight = 13
  object 日志记录: TLabel
    Left = 144
    Top = 24
    Width = 48
    Height = 13
    Caption = #26085#24535#35760#24405
  end
  object Label1: TLabel
    Left = 144
    Top = 204
    Width = 48
    Height = 13
    Caption = #28165#29702#32467#26524
  end
  object Label2: TLabel
    Left = 144
    Top = 408
    Width = 252
    Height = 13
    Caption = #22914#26524#35201#25163#21160#20462#22797#65292#35831#22312#27492#36755#20837#35201#24674#22797#30340#25991#20214#36335#24452
  end
  object Button1: TButton
    Left = 24
    Top = 192
    Width = 75
    Height = 25
    Caption = #24555#36895#20462#22797
    TabOrder = 0
    OnClick = Button1Click
  end
  object ListBox2: TListBox
    Left = 144
    Top = 232
    Width = 625
    Height = 148
    ItemHeight = 13
    TabOrder = 1
  end
  object ListBox1: TListBox
    Left = 144
    Top = 56
    Width = 625
    Height = 129
    ItemHeight = 13
    TabOrder = 2
  end
  object MaskEdit1: TMaskEdit
    Left = 144
    Top = 440
    Width = 625
    Height = 21
    TabOrder = 3
    Text = ''
  end
  object Button2: TButton
    Left = 24
    Top = 435
    Width = 75
    Height = 25
    Caption = #25163#21160#26597#26432#25991#20214
    TabOrder = 4
    OnClick = Button2Click
  end
end
