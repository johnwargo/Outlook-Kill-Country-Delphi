object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = 'Empty All Outlook Contact Country Fields'
  ClientHeight = 461
  ClientWidth = 738
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -16
  Font.Name = 'Hack'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 19
  object StatusBar1: TStatusBar
    Left = 0
    Top = 442
    Width = 738
    Height = 19
    Panels = <>
    SimplePanel = True
    SimpleText = 'By John M. Wargo (www.johnwargo.com)'
  end
  object output: TMemo
    Left = 0
    Top = 0
    Width = 738
    Height = 442
    Align = alClient
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
  end
end
