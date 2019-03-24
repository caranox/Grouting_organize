object Form1: TForm1
  Left = 505
  Top = 378
  Width = 700
  Height = 550
  Caption = #28748#27974#25968#25454#32479#35745#25972#29702
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    684
    512)
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 596
    Top = 0
    Width = 88
    Height = 412
    Anchors = [akTop, akRight, akBottom]
    TabOrder = 0
    DesignSize = (
      88
      412)
    object ListBox1: TListBox
      Left = 3
      Top = 165
      Width = 80
      Height = 244
      Anchors = [akLeft, akTop, akBottom]
      ItemHeight = 13
      TabOrder = 0
      OnClick = ListBox1Click
    end
    object Button1: TButton
      Left = 3
      Top = 59
      Width = 80
      Height = 25
      Caption = #35835#21462'Excel'
      TabOrder = 1
      OnClick = Button1Click
    end
    object Button4: TButton
      Left = 3
      Top = 85
      Width = 80
      Height = 25
      Caption = #25209#37327#23548#20837
      TabOrder = 2
      OnClick = Button4Click
    end
    object Button2: TButton
      Left = 3
      Top = 32
      Width = 80
      Height = 25
      Caption = #38142#25509#25968#25454#24211
      TabOrder = 3
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 3
      Top = 5
      Width = 80
      Height = 25
      Caption = #26032#24314#25968#25454#24211
      TabOrder = 4
      OnClick = Button3Click
    end
    object Button5: TButton
      Left = 3
      Top = 138
      Width = 80
      Height = 25
      Caption = #32479#35745#25968#25454
      TabOrder = 5
      OnClick = Button5Click
    end
    object Button6: TButton
      Left = 3
      Top = 112
      Width = 80
      Height = 25
      Caption = #25972#29702#25968#25454
      TabOrder = 6
      OnClick = Button6Click
    end
  end
  object DbgEh_S: TDBGridEh
    Left = 0
    Top = 0
    Width = 597
    Height = 161
    Color = clInactiveCaption
    DynProps = <>
    FooterParams.Color = clWindow
    TabOrder = 1
    OnDblClick = DbgEh_SDblClick
    object RowDetailData: TRowDetailPanelControlEh
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 493
    Width = 684
    Height = 19
    AutoHint = True
    Panels = <
      item
        Width = 200
      end
      item
        Width = 200
      end
      item
        Width = 200
      end>
    OnClick = StatusBar1Click
  end
  object ListBox2: TListBox
    Left = 0
    Top = 415
    Width = 392
    Height = 80
    Anchors = [akLeft, akRight, akBottom]
    ItemHeight = 13
    TabOrder = 3
    OnClick = ListBox2Click
    OnDblClick = ListBox2DblClick
  end
  object DBChart1: TDBChart
    Left = -88
    Top = 0
    Width = 601
    Height = 323
    AutoRefresh = False
    ShowGlassCursor = False
    BackWall.Brush.Color = clWhite
    Title.Text.Strings = (
      'TDBChart')
    LeftAxis.TickLength = 2
    Legend.ColorWidth = 20
    Legend.Inverted = True
    Legend.LegendStyle = lsValues
    Legend.ShadowSize = 0
    Legend.TextStyle = ltsLeftPercent
    Legend.TopPos = 50
    Legend.Visible = False
    View3D = False
    View3DOptions.Elevation = 315
    View3DOptions.Orthogonal = False
    View3DOptions.Perspective = 0
    View3DOptions.Rotation = 360
    ParentShowHint = False
    ShowHint = False
    TabOrder = 4
    Anchors = [akLeft, akTop, akRight, akBottom]
    OnDblClick = DBChart1DblClick
    object Series1: TPointSeries
      ColorEachPoint = True
      Marks.ArrowLength = 1
      Marks.Frame.Visible = False
      Marks.Transparent = True
      Marks.Visible = False
      SeriesColor = clRed
      Pointer.InflateMargins = True
      Pointer.Style = psRectangle
      Pointer.Visible = True
      XValues.DateTime = False
      XValues.Name = 'X'
      XValues.Multiplier = 1.000000000000000000
      XValues.Order = loAscending
      YValues.DateTime = False
      YValues.Name = 'Y'
      YValues.Multiplier = 1.000000000000000000
      YValues.Order = loNone
    end
  end
  object Memo1: TMemo
    Left = 392
    Top = 415
    Width = 145
    Height = 80
    Anchors = [akRight, akBottom]
    ScrollBars = ssBoth
    TabOrder = 5
  end
  object Memo2: TMemo
    Left = 539
    Top = 415
    Width = 145
    Height = 80
    Anchors = [akRight, akBottom]
    ScrollBars = ssBoth
    TabOrder = 6
  end
  object DS_S: TDataSource
    Left = 152
    Top = 296
  end
  object Qur_S: TADOQuery
    Parameters = <>
    Left = 192
    Top = 296
  end
  object ADOC_S: TADOConnection
    Left = 112
    Top = 296
  end
  object OpenDialog1: TOpenDialog
    Left = 32
    Top = 296
  end
  object SaveDialog1: TSaveDialog
    Left = 72
    Top = 296
  end
  object PopupMenu1: TPopupMenu
    Left = 144
    Top = 64
    object N1: TMenuItem
      Caption = #25171#21360#34920#26684
      OnClick = N1Click
    end
    object XT1: TMenuItem
      Caption = #23548#20986#25968#25454
      OnClick = XT1Click
    end
    object N5: TMenuItem
      Caption = #38544#34255#22270
      OnClick = N5Click
    end
    object N2: TMenuItem
      Caption = #23380#28145'-'#21525#33635#22270
      OnClick = N2Click
    end
    object N3: TMenuItem
      Caption = #23380#28145'-'#28748#27974#37327#22270
      OnClick = N3Click
    end
    object N4: TMenuItem
      Caption = #21525#33635'-'#28748#27974#37327#22270
      OnClick = N4Click
    end
    object N6: TMenuItem
      Caption = #25968#25454#20462#25913
      OnClick = N6Click
    end
  end
  object ADOQuery2: TADOQuery
    Parameters = <>
    Left = 192
    Top = 272
  end
  object ADOConnection2: TADOConnection
    Left = 112
    Top = 272
  end
  object DataSource2: TDataSource
    Left = 152
    Top = 272
  end
  object PrintDBGridEh2: TPrintDBGridEh
    Options = []
    PageFooter.Font.Charset = DEFAULT_CHARSET
    PageFooter.Font.Color = clWindowText
    PageFooter.Font.Height = -11
    PageFooter.Font.Name = 'MS Sans Serif'
    PageFooter.Font.Style = []
    PageHeader.Font.Charset = DEFAULT_CHARSET
    PageHeader.Font.Color = clWindowText
    PageHeader.Font.Height = -11
    PageHeader.Font.Name = 'MS Sans Serif'
    PageHeader.Font.Style = []
    Units = MM
    Left = 224
    Top = 288
  end
end
