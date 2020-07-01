VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _Version        =   786432
      _ExtentX        =   18865
      _ExtentY        =   9340
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "File lock          "
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "File unlock            "
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   4695
         Left            =   30
         TabIndex        =   2
         Top             =   570
         Width           =   10635
         _Version        =   786432
         _ExtentX        =   18759
         _ExtentY        =   8281
         _StockProps     =   1
         Page            =   1
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2760
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   3840
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   360
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   3840
            Visible         =   0   'False
            Width           =   1935
         End
         Begin XtremeSuiteControls.PushButton PushButton4 
            Height          =   855
            Left            =   3120
            TabIndex        =   16
            Top             =   3240
            Width           =   3375
            _Version        =   786432
            _ExtentX        =   5953
            _ExtentY        =   1508
            _StockProps     =   79
            Caption         =   "Unlock now"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Checked         =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit3 
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   2400
            Width           =   6975
            _Version        =   786432
            _ExtentX        =   12303
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   495
            Left            =   7680
            TabIndex        =   13
            Top             =   840
            Width           =   2415
            _Version        =   786432
            _ExtentX        =   4260
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Select file you want to unlock"
            UseVisualStyle  =   -1  'True
         End
         Begin RichTextLib.RichTextBox rr 
            Height          =   1575
            Left            =   9840
            TabIndex        =   12
            Top             =   2040
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   2778
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"FILE PROTECTOR.frx":0000
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   873
            _StockProps     =   79
            Transparent     =   -1  'True
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   4695
         Left            =   -69970
         TabIndex        =   1
         Top             =   570
         Visible         =   0   'False
         Width           =   10635
         _Version        =   786432
         _ExtentX        =   18759
         _ExtentY        =   8281
         _StockProps     =   1
         Page            =   0
         Begin RichTextLib.RichTextBox R 
            Height          =   1815
            Left            =   7680
            TabIndex        =   11
            Top             =   2880
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   3201
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"FILE PROTECTOR.frx":008B
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   1455
            Left            =   7920
            TabIndex        =   10
            Top             =   2160
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   2566
            _StockProps     =   79
            Caption         =   "LOCK FILE NOW"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Checked         =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   2895
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Width           =   7215
            _Version        =   786432
            _ExtentX        =   12726
            _ExtentY        =   5106
            _StockProps     =   79
            Transparent     =   -1  'True
            Appearance      =   5
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   375
               Left            =   120
               TabIndex        =   7
               Top             =   1800
               Width           =   6495
               _Version        =   786432
               _ExtentX        =   11456
               _ExtentY        =   661
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   120
               TabIndex        =   6
               Top             =   720
               Width           =   6495
               _Version        =   786432
               _ExtentX        =   11456
               _ExtentY        =   661
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Left            =   360
               TabIndex        =   9
               Top             =   1440
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Confrom your password"
               ForeColor       =   255
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   195
               Left            =   360
               TabIndex        =   8
               Top             =   360
               Width           =   1440
               _Version        =   786432
               _ExtentX        =   2540
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Enter your password"
               ForeColor       =   255
               Transparent     =   -1  'True
               AutoSize        =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit Fl 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   7335
            _Version        =   786432
            _ExtentX        =   12938
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   375
            Left            =   7800
            TabIndex        =   3
            Top             =   720
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Select a file"
            BackColor       =   -2147483635
            Appearance      =   6
            Checked         =   -1  'True
         End
         Begin XtremeSuiteControls.CommonDialog CommonDialog1 
            Left            =   7800
            Top             =   1440
            _Version        =   786432
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   4
         End
      End
   End
   Begin XtremeSuiteControls.CommonDialog CommonDialog2 
      Left            =   4560
      Top             =   2280
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatEdit2_Change()
If FlatEdit1.Text = FlatEdit2.Text And FlatEdit1.Text <> "" And FlatEdit2.Text <> "" Then
PushButton2.Enabled = True
FlatEdit1.Enabled = False
FlatEdit2.Enabled = False

End If

End Sub

Private Sub PushButton1_Click()
FlatEdit1.Enabled = True
FlatEdit2.Enabled = True

CommonDialog1.ShowOpen
Fl.Text = CommonDialog1.FileName
R.LoadFile Fl.Text



End Sub

Private Sub PushButton2_Click()
Dim fso As New FileSystemObject
Dim f As File
Dim t As TextStream
Set f = fso.GetFile(Fl.Text)
R.Text = "kjdfgyutdvchjcv/jsdfpog]a[]jo][][ou6ktljhg;lgjkbvrtuyipojb hpoe4u5ytgjkhnoiuyjtg;bnklj'j{}ptyopkplukltedkyjkykljuklpupjkouyjkpoyj" & FlatEdit1.Text & vbCrLf & f.Name & vbCrLf & R.Text
Set t = fso.CreateTextFile(f.ParentFolder & "\" & f.ShortName & ".exe")
t.Write R.Text
t.Close
Kill Fl.Text
FlatEdit1.Text = ""
FlatEdit2.Text = ""
Fl.Text = ""
PushButton2.Enabled = False

End Sub

Private Sub PushButton3_Click()
On Error GoTo ss:
Dim fso As New FileSystemObject
Dim f As File
Dim t As TextStream

CommonDialog2.ShowOpen
Label3.Caption = CommonDialog2.FileName
rr.LoadFile Label3.Caption
Set t = fso.OpenTextFile(Label3.Caption, ForReading)
Text1.Text = t.ReadLine
Text2.Text = t.ReadLine
rr.Find Text1.Text
rr.SelText = ""
rr.Find Text2.Text
rr.SelText = ""
rr.Find vbCrLf & vbCrLf
rr.SelText = ""
t.Close
FlatEdit3.Enabled = True
PushButton4.Enabled = True
ss:
End Sub

Private Sub PushButton4_Click()
Dim a As String
Dim w As Integer
w = Len(Text1.Text)
a = Mid$(Text1.Text, 130, w - 129)
If a = FlatEdit3.Text Then
Dim fso As New FileSystemObject
Dim f As File
Dim t As TextStream
Set f = fso.GetFile(Label3.Caption)
Set t = fso.CreateTextFile(f.ParentFolder & "\" & Text2.Text)
t.Write rr.Text
t.Close
FlatEdit3.Enabled = False
PushButton4.Enabled = False
Kill f.Path
MsgBox "Unlock complete", vbOKOnly, "Complete"
FlatEdit3.Text = ""
Else
MsgBox "Try again", vbCritical, "Your password is wrong"
End If
End Sub
