VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin Project1.PropertyBox PropertyBox1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _extentx        =   8705
      _extenty        =   4683
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    PropertyBox1.CheckboxVisible = Not PropertyBox1.CheckboxVisible
    
End Sub

Private Sub Command2_Click()
   PropertyBox1.Locked = Not PropertyBox1.Locked
End Sub

Private Sub Form_Load()

    PropertyBox1.AddProperty "Appearance", "COMBO", , "|0 - Flat|1 - 3D|"
    PropertyBox1.AddProperty "AutoRedraw", "COMBO", , "|False|True|"
    PropertyBox1.AddProperty "BorderStyle", "COMBO", , "|0 - None|1 - Fixed Single|2 - Fixed Dialog|"
    PropertyBox1.AddProperty "Caption", "TEXT"
    PropertyBox1.AddProperty "Clip Controls", "COMBO", , "|False|True|"
    PropertyBox1.AddProperty "DrawMode", "TEXT"
    PropertyBox1.AddProperty "DrawStyle", "TEXT"
    PropertyBox1.AddProperty "DrawWidth", "TEXT"
    PropertyBox1.AddProperty "Enabled", "TEXT"
    PropertyBox1.AddProperty "Font", "BUTTON"
    PropertyBox1.AddProperty "FontTransparent", "BUTTON"
    PropertyBox1.DrawPropertyBox
    PropertyBox1.CheckboxVisible = False
  
''''''    PropertyBox1.Locked = True
  
End Sub





