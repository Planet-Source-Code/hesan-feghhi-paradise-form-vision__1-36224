VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BackColor       =   &H009E7000&
   Caption         =   "Paradise Form Vision"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   Icon            =   "Form Vision.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox TabContent 
      BackColor       =   &H00D7A953&
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   2
      Left            =   3300
      ScaleHeight     =   7455
      ScaleWidth      =   10335
      TabIndex        =   2
      Top             =   3120
      Width           =   10335
      Begin VB.Shape Shape7 
         Height          =   735
         Left            =   1500
         Top             =   2580
         Width           =   7395
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Form Vision.frx":030A
         Height          =   615
         Left            =   1620
         TabIndex        =   67
         Top             =   2640
         Width           =   7215
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F4B82B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7800
         TabIndex        =   63
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label25 
         BackColor       =   &H00F8D074&
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008C6E00&
         Height          =   855
         Left            =   240
         TabIndex        =   62
         Top             =   4680
         Width           =   9855
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Never use this item on a programmed form. It will remove all the existing form contents."
         Height          =   195
         Left            =   2820
         TabIndex        =   61
         Top             =   4140
         Width           =   6135
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "WARNING:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1620
         TabIndex        =   60
         Top             =   4140
         Width           =   1155
      End
      Begin VB.Label Label20 
         BackColor       =   &H00F8D074&
         Caption         =   " Creation "
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008C6E00&
         Height          =   735
         Left            =   240
         TabIndex        =   59
         Top             =   1080
         Width           =   9855
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FAE1A7&
         FillStyle       =   0  'Solid
         Height          =   2895
         Left            =   240
         Top             =   1800
         Width           =   9855
      End
   End
   Begin VB.PictureBox TabContent 
      BackColor       =   &H00D7A953&
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   1
      Left            =   6000
      ScaleHeight     =   7455
      ScaleWidth      =   10335
      TabIndex        =   1
      Top             =   3360
      Width           =   10335
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   3240
         ScaleHeight     =   2625
         ScaleWidth      =   3465
         TabIndex        =   34
         Top             =   2040
         Width           =   3495
         Begin MSForms.TextBox TopTB 
            Height          =   255
            Left            =   1680
            TabIndex        =   52
            Top             =   2160
            Width           =   1815
            VariousPropertyBits=   746604571
            BackColor       =   16306292
            Size            =   "3201;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox CheckBox5 
            Height          =   255
            Left            =   2400
            TabIndex        =   50
            Top             =   960
            Width           =   255
            VariousPropertyBits=   746588179
            BackColor       =   10383360
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "450;450"
            Value           =   "1"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox CheckBox4 
            Height          =   255
            Left            =   2400
            TabIndex        =   49
            Top             =   960
            Width           =   255
            VariousPropertyBits=   746588179
            BackColor       =   10383360
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "450;450"
            Value           =   "1"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox CheckBox3 
            Height          =   255
            Left            =   2400
            TabIndex        =   41
            Top             =   1920
            Width           =   255
            VariousPropertyBits=   746588179
            BackColor       =   10383360
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "450;450"
            Value           =   "1"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox CheckBox2 
            Height          =   255
            Left            =   2400
            TabIndex        =   39
            Top             =   720
            Width           =   255
            VariousPropertyBits=   746588179
            BackColor       =   10383360
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "450;450"
            Value           =   "1"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox CheckBox1 
            Height          =   255
            Left            =   2400
            TabIndex        =   36
            Top             =   480
            Width           =   255
            VariousPropertyBits=   746588179
            BackColor       =   10383360
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "450;450"
            Value           =   "1"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox WidthTB 
            Height          =   255
            Left            =   1680
            TabIndex        =   56
            Top             =   2400
            Width           =   1815
            VariousPropertyBits=   746604571
            BackColor       =   16037931
            Size            =   "3201;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00F4B82B&
            FillColor       =   &H00F4B82B&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   2
            Left            =   1680
            Top             =   1920
            Width           =   1815
         End
         Begin MSForms.TextBox HeightTB 
            Height          =   255
            Left            =   1680
            TabIndex        =   54
            Top             =   1440
            Width           =   1815
            VariousPropertyBits=   746604571
            BackColor       =   16037931
            Size            =   "3201;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox LeftTB 
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   1200
            Width           =   1815
            VariousPropertyBits=   746604571
            BackColor       =   16306292
            Size            =   "3201;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00F4B82B&
            FillColor       =   &H00F4B82B&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   3
            Left            =   1680
            Top             =   960
            Width           =   1815
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00F8D074&
            FillColor       =   &H00F8D074&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   1
            Left            =   1680
            Top             =   720
            Width           =   1815
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00F4B82B&
            FillColor       =   &H00F4B82B&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   0
            Left            =   1680
            Top             =   480
            Width           =   1815
         End
         Begin MSForms.ComboBox CB1 
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   240
            Width           =   1815
            VariousPropertyBits=   614483995
            BackColor       =   16306292
            DisplayStyle    =   7
            Size            =   "3201;450"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox NameTB 
            Height          =   255
            Left            =   1695
            TabIndex        =   47
            Top             =   0
            Width           =   1800
            VariousPropertyBits=   746604571
            BackColor       =   16037931
            Size            =   "3175;450"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label11 
            BackColor       =   &H00F8D074&
            Caption         =   " Width"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   55
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FAE1A7&
            Caption         =   " Top"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   51
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackColor       =   &H00F8D074&
            Caption         =   " Height"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   53
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FAE1A7&
            Caption         =   " Left"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   44
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label18 
            BackColor       =   &H00F8D074&
            Caption         =   " Control Box"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   48
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FAE1A7&
            Caption         =   " Clip Controls"
            Height          =   255
            Left            =   0
            TabIndex        =   40
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H00F8D074&
            Caption         =   " Auto Redraw"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   38
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FAE1A7&
            Caption         =   " Appearance"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H00F8D074&
            Caption         =   " (Name)"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label18 
            BackColor       =   &H00F8D074&
            Caption         =   " Show in Taskbar"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   42
            Top             =   1920
            Width           =   1695
         End
         Begin MSForms.ComboBox CB2 
            Height          =   255
            Left            =   1680
            TabIndex        =   58
            Top             =   1680
            Width           =   1815
            VariousPropertyBits=   614483995
            BackColor       =   16306292
            DisplayStyle    =   7
            Size            =   "3201;450"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FAE1A7&
            Caption         =   " Shape"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   57
            Top             =   1680
            Width           =   1695
         End
      End
      Begin VB.Shape Shape3 
         Height          =   3255
         Left            =   3120
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E7A418&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Form Properties"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   43
         Top             =   1680
         Width           =   3495
      End
   End
   Begin VB.PictureBox TabContent 
      BackColor       =   &H00D7A953&
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   0
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   10335
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame1 
         BackColor       =   &H00D7A953&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3960
         TabIndex        =   64
         Top             =   360
         Width           =   5895
         Begin VB.Image Sensor 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1920
            Top             =   0
            Width           =   3015
         End
         Begin VB.Shape Shape6 
            Height          =   255
            Index           =   1
            Left            =   2520
            Top             =   0
            Width           =   615
         End
         Begin VB.Shape Shape6 
            Height          =   255
            Index           =   3
            Left            =   3720
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Please Vote on this Code:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00BF8B00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Go to PSC"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5040
            TabIndex        =   65
            Top             =   0
            Width           =   855
         End
         Begin VB.Shape Rating 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   1920
            Top             =   0
            Width           =   2175
         End
         Begin VB.Shape Shape8 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   255
            Left            =   1920
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   3600
         Top             =   2760
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FAE1A7&
         Height          =   285
         Left            =   840
         TabIndex        =   29
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FAE1A7&
         Height          =   285
         Left            =   2280
         TabIndex        =   28
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FAE1A7&
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Top             =   2280
         Width           =   375
      End
      Begin VB.PictureBox DSC 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   3960
         ScaleHeight     =   5145
         ScaleWidth      =   5865
         TabIndex        =   25
         Top             =   960
         Width           =   5895
         Begin VB.PictureBox DS 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00E7A418&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5175
            Left            =   0
            ScaleHeight     =   5175
            ScaleWidth      =   5895
            TabIndex        =   26
            Top             =   0
            Width           =   5895
            Begin VB.Shape BB 
               BorderStyle     =   3  'Dot
               Height          =   1095
               Left            =   1440
               Top             =   2760
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.Line PL 
               BorderColor     =   &H000000FF&
               Visible         =   0   'False
               X1              =   3600
               X2              =   1920
               Y1              =   480
               Y2              =   2160
            End
         End
      End
      Begin MSComDlg.CommonDialog CDlg 
         Left            =   9840
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Point Selection:"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1920
         TabIndex        =   32
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clone"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   31
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "     X:                            Y:"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H009E7000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Use the Arrow Keys to Move around the Background"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   24
         Top             =   6120
         Width           =   5895
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7A418&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Form Shape and Background View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9600
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   2040
         TabIndex        =   21
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   2640
         TabIndex        =   20
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7A418&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Spline Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Union"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   2160
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   3360
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3480
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Spline Overlay Method :"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Spline Selection:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">>"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   2640
         TabIndex        =   14
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   2040
         TabIndex        =   13
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   12
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "< < "
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add Spline"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove Spline"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00BC750A&
         FillStyle       =   0  'Solid
         Height          =   4935
         Left            =   240
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7A418&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Form Shape"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove Picture"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BF8B00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Open Picture"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7A418&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00BC750A&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   240
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.Image TabMO 
      Height          =   1215
      Index           =   2
      Left            =   10320
      ToolTipText     =   "Creation"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image TabMO 
      Height          =   1095
      Index           =   1
      Left            =   10320
      ToolTipText     =   "Properties"
      Top             =   960
      Width           =   255
   End
   Begin VB.Image TabPic 
      Height          =   255
      Index           =   2
      Left            =   10320
      Picture         =   "Form Vision.frx":0412
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   255
   End
   Begin VB.Image TabPic 
      Height          =   255
      Index           =   1
      Left            =   10320
      Picture         =   "Form Vision.frx":2E86
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   255
   End
   Begin VB.Image TabPic 
      Height          =   255
      Index           =   0
      Left            =   10320
      Picture         =   "Form Vision.frx":58FA
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   255
   End
   Begin VB.Image TabMO 
      Height          =   975
      Index           =   0
      Left            =   10320
      ToolTipText     =   "Design"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image TabImg 
      Height          =   7500
      Left            =   10320
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VoteHeld

Private Type PointXY
 X As Integer
 Y As Integer
End Type

Private Points(55, 401) As PointXY
Private OvlMethod(55) As Integer
Private CurPG, CurP, Started_Create As Boolean, Started_Move As Boolean, xx, yy, SelPG, SelP(55), Started_Move_Whole As Boolean
Private OMS(4)

Private Sub Tabs_Click(Index As Integer)
 TabBar.BackColor = Tabs(Index).BackColor
End Sub

Private Sub DS_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 33
   Label2_Click 5
  Case 34
   Label2_Click 6
  Case 35
   Label2_Click 7
  Case 36
   Label2_Click 4
  Case 37
   If DS.Left < 0 Then DS.Left = DS.Left + 150
   If DS.Left > 0 Then DS.Left = 0
  Case 38
   If DS.Top < 0 Then DS.Top = DS.Top + 150
   If DS.Top > 0 Then DS.Top = 0
  Case 39
   If DS.Left + DS.Width > DSC.Width Then DS.Left = DS.Left - 150
   If DS.Left + DS.Width < DSC.Width Then DS.Left = DSC.Width - DS.Width
  Case 40
   If DS.Top + DS.Height > DSC.Height Then DS.Top = DS.Top - 150
   If DS.Top + DS.Height < DSC.Height Then DS.Top = DSC.Height - DS.Height
  Case 45
   Label2_Click 2
  Case 46
   Label2_Click 3
  Case 107
   Label2_Click 9
  Case 109
   Label2_Click 10
 End Select
End Sub

Private Sub DS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If CurPG = -1 Then CurPG = 0
 If Started_Create = True Then
  If Button = 1 Then
   Points(CurPG, CurP).X = X
   Points(CurPG, CurP).Y = Y
   PL.Visible = True
   PL.X1 = X
   PL.Y1 = Y
   If CurP = 0 Then
    DS.Line (X - 30, Y - 30)-(X + 30, Y + 30), RGB(0, 255, 0), B
    DS.PSet (X, Y)
   Else
    DS.Line -(X, Y), RGB(255, 0, 0)
    DS.Line (X - 30, Y - 30)-(X + 30, Y + 30), RGB(0, 255, 0), B
    DS.PSet (X, Y)
   End If
   CurP = CurP + 1
  ElseIf Button = 2 Then
   PL.Visible = False
   DS.Line -(Points(CurPG, 0).X, Points(CurPG, 0).Y), RGB(255, 0, 0)
   Points(CurPG, CurP).X = Points(CurPG, 0).X
   Points(CurPG, CurP).Y = Points(CurPG, 0).Y
   CurP = CurP + 1
   Points(CurPG, CurP).X = -1
   Started_Create = False
   SelPG = CurPG
   CurP = 0
   SelP(SelPG) = 1
   CurPG = CurPG + 1
   RedrawAll
  End If
 Else
  If CurPG = 0 Or SelPG = -1 Then Exit Sub
  Started_Move = True
  SelP(SelPG) = -1
  For i = 0 To 400
   If Points(SelPG, i).X = -1 Then Exit For
   px = Points(SelPG, i).X
   py = Points(SelPG, i).Y
   If X > px - 50 And X < px + 50 And Y > py - 50 And Y < py + 50 Then
    SelP(SelPG) = i
    xx = X
    yy = Y
   End If
  Next i
  If SelP(SelPG) = -1 Then Started_Move = False: Started_Move_Whole = True: xx = X: yy = Y
  RedrawAll
 End If
End Sub

Private Sub DS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 PL.X2 = X
 PL.Y2 = Y
End Sub

Private Sub DS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Started_Move = True Then
  Points(SelPG, SelP(SelPG)).X = X
  Points(SelPG, SelP(SelPG)).Y = Y
  If Points(SelPG, SelP(SelPG) + 1).X = -1 Then
   Points(SelPG, 0).X = X
   Points(SelPG, 0).Y = Y
  End If
  RedrawAll
  Started_Move = False
 End If
 If Started_Move_Whole = True Then
  For i = 0 To 400
   If Points(SelPG, i).X = -1 Then Exit For
   Points(SelPG, i).X = Points(SelPG, i).X + X - xx
   Points(SelPG, i).Y = Points(SelPG, i).Y + Y - yy
  Next i
  RedrawAll
  Started_Move_Whole = False
 End If
End Sub

Private Sub Form_Load()
 TabImg.Picture = TabPic(0).Picture
 For i = 0 To 2
  TabContent(i).Left = 0
  TabContent(i).Top = 0
  TabContent(i).Visible = False
 Next i
 TabContent(0).Visible = True
 For i = 1 To 55
  Points(i, 0).X = -1
  OvlMethod(i) = 0
 Next i
 OMS(0) = "Union"
 OMS(1) = "Copy"
 OMS(2) = "Intersection"
 OMS(3) = "Substraction"
 OMS(4) = "XOR Union"
 CB1.AddItem "3D"
 CB1.AddItem "Flat"
 CB2.AddItem "Automatic"
 CB2.AddItem "Manual"
 CB1.ListIndex = 0
 CB2.ListIndex = 1
 LeftTB = 0
 TopTB = 0
 WidthTB = 3000
 HeightTB = 2200
 NameTB = "Form1"
End Sub

Private Sub Label16_Click()
 MsgBox "This part is for advanced users. On this part, you can parametrically set the vertices' position, or do some actions with vertices. By cloning, the selected vertex (shown in blue) will divide into 2 seperate ones. By removing, the vertex will be deleted and a line +will connect the vertexes before and after the deleted one.", vbInformation, "Help"
End Sub

Private Sub Label2_Click(Index As Integer)
 Select Case Index
  Case 0
   'Open the Background Picture
   CDlg.FileName = ""
   CDlg.DialogTitle = "Open a Background Picture..."
   CDlg.Filter = "Bitmap Files (*.bmp)|*.bmp|Jpeg Files (*.jpg;*.jpe)|*.jpg;*.jpe|GIF Files (*.gif)|*.gif|Icons and Cursors (*.ico;*.cur)|*.ico;*.cur|All Picture Files (*.bmp;*.jpg;*.jpe;*.gif;*.ico;*.cur)|*.bmp;*.jpg;*.jpe;*.gif;*.ico;*.cur"
   CDlg.ShowOpen
   If CDlg.FileName <> "" Then
    DS.Picture = LoadPicture(CDlg.FileName)
    DS.Tag = CDlg.FileName
   End If
   RedrawAll
  Case 1
   'Remove the Background Picture
   DS.Picture = LoadPicture("")
   DS.Tag = ""
   RedrawAll
  Case 2
   'Start the Spline Creation
   Started_Create = True
  Case 3
   If SelPG = -1 Then Exit Sub
   For j = SelPG To 55
    For i = 0 To 490
     Points(j, i) = Points(j + 1, i)
     If Points(j, i).X = -1 Then Exit For
1   Next i
    If Points(j, 0).X = -1 Then Exit For
   Next j
   BB.Visible = False
   SelPG = -1
   RedrawAll
   CurPG = CurPG - 1
  Case 4
   SelPG = 0
   RedrawAll
  Case 5
   If SelPG > 0 Then SelPG = SelPG - 1
   RedrawAll
  Case 6
   If Points(SelPG + 1, 0).X > -1 Then SelPG = SelPG + 1
   RedrawAll
  Case 7
   For i = 0 To 55
    If Points(i, 0).X = -1 Then SelPG = i - 1: Exit For
   Next i
   RedrawAll
  Case 8
   OvlMethod(SelPG) = OvlMethod(SelPG) + 1
   If OvlMethod(SelPG) = -1 Then OvlMethod(SelPG) = 4
   If OvlMethod(SelPG) = 5 Then OvlMethod(SelPG) = 0
   Label2(8).Caption = OMS(OvlMethod(SelPG))
  Case 9
   If SelP(SelPG) = -1 Then Exit Sub
   If SelP(SelPG) = 0 Or Points(SelPG, SelP(SelPG) + 1).X = -1 Then MsgBox "You cannot clone or remove the first and the last point of the polygon": Exit Sub
   For i = 400 To SelP(SelPG) Step -1
    Points(SelPG, i).X = Points(SelPG, i - 1).X
    Points(SelPG, i).Y = Points(SelPG, i - 1).Y
   Next i
   Points(SelPG, SelP(SelPG)).X = Points(SelPG, SelP(SelPG) + 1).X + 45
   Points(SelPG, SelP(SelPG)).Y = Points(SelPG, SelP(SelPG) + 1).Y + 45
   For i = SelP(SelPG) To 400
    If Points(SelPG, i).X = -1 Then
     Points(SelPG, 0).X = Points(SelPG, i - 1).X
     Points(SelPG, 0).Y = Points(SelPG, i - 1).Y
     Exit For
    End If
   Next i
   RedrawAll
  Case 10
   If SelP(SelPG) = -1 Then Exit Sub
   If SelP(SelPG) = 0 Or Points(SelPG, SelP(SelPG) + 1).X = -1 Then MsgBox "You cannot clone or remove the first and the last point of the polygon": Exit Sub
   For i = SelP(SelPG) To 400
    Points(SelPG, i).X = Points(SelPG, i + 1).X
    Points(SelPG, i).Y = Points(SelPG, i + 1).Y
    If Points(SelPG, i).X = -1 Then
     Points(SelPG, 0).X = Points(SelPG, i - 1).X
     Points(SelPG, 0).Y = Points(SelPG, i - 1).Y
     Exit For
    End If
   Next i
   RedrawAll
  Case 11
   SelP(SelPG) = 0
   RedrawAll
  Case 12
   If SelP(SelPG) > 0 Then SelP(SelPG) = SelP(SelPG) - 1
   RedrawAll
  Case 13
   If Points(SelPG, SelP(SelPG) + 1).X > -1 Then SelP(SelPG) = SelP(SelPG) + 1 Else SelP(SelPG) = 1
   RedrawAll
  Case 14
   For i = 0 To 400
    If Points(SelPG, i).X = -1 Then SelP(SelPG) = i - 1: Exit For
   Next i
   RedrawAll
 End Select
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2(Index).BackColor = RGB(0, 150, 255)
End Sub

Private Sub Label2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2(Index).BackColor = &HBF8B00
End Sub

Private Sub Label26_Click()
1 CDlg.FileName = ""
  CDlg.DialogTitle = "Open Visual Basic Form File ..."
  CDlg.Filter = "Visual Basic Form Files (*.frm)|*.frm"
  CDlg.ShowSave
  If CDlg.FileName <> "" Then
   If FileExists(CDlg.FileName) Then
    a = MsgBox("This File Already Exists. Do you want to replace it?", vbYesNoCancel, "Save Form File")
    If a = vbYes Then
     GoTo 2
    ElseIf a = vbNo Then GoTo 1
    Else: Exit Sub
    End If
   Else
2  'Build The Form File
    Call BuildFormFile(CDlg.FileName)
   End If
  End If
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label26.BackColor = &HF4B88B
End Sub

Private Sub Label26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label26.BackColor = &HF4B82B
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label28.BackColor = RGB(200, 100, 0)
End Sub

Private Sub Label28_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label28.BackColor = &H9E7000
 frmVote.WB.Navigate "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=36224"
 frmVote.Show vbModal
End Sub

Private Sub Label3_Click()
 MsgBox "Click on 'Add Spline' and then left-click on this page to add vertices and when you are done, right-click to close the spline. To move any vertex, just drag it. To move a whole spline, drag on a non-vertex area of the view, to change the spline pointer, use the 'Spline Selection' tools on the left, to move around the background, use Arrow Keys.", vbInformation, "Help"
End Sub

Private Sub Label4_Click()
 MsgBox "Use these buttons to open or clear the background picture", vbInformation, "Help"
End Sub

Private Sub Label6_Click()
 MsgBox "Using this part, you can control Splines, Add or Remove them. For more information click on the preview box's Help Button ('?').", vbInformation, "Help"
End Sub

Private Sub Label8_Click()
 DS.SetFocus
End Sub

Private Sub LeftTB_Change()
 Static R$
 If StrInText("qwertyuiop[]}{POIUYTREWQasdfghjkl;'"":LKJHGFDSAzxcvbnm,./?><MNBVCXZ`~!@#$%^&*()_+=-|\", LeftTB) Then
  MsgBox "Invalid property value : " + LeftTB, vbCritical, "Error"
  LeftTB = R$
 End If
 R$ = LeftTB
End Sub

Private Sub TopTB_Change()
 Static R$
 If StrInText("qwertyuiop[]}{POIUYTREWQasdfghjkl;'"":LKJHGFDSAzxcvbnm,./?><MNBVCXZ`~!@#$%^&*()_+=-|\", TopTB) Then
  MsgBox "Invalid property value : " + TopTB, vbCritical, "Error"
  TopTB = R$
 End If
 R$ = TopTB
End Sub

Private Sub WidthTB_Change()
 Static R$
 If StrInText("qwertyuiop[]}{POIUYTREWQasdfghjkl;'"":LKJHGFDSAzxcvbnm,./?><MNBVCXZ`~!@#$%^&*()_+=-|\", WidthTB) Then
  MsgBox "Invalid property value : " + WidthTB, vbCritical, "Error"
  WisthTB = R$
 End If
 R$ = WidthTB
End Sub

Private Sub HeightTB_Change()
 Static R$
 If StrInText("qwertyuiop[]}{POIUYTREWQasdfghjkl;'"":LKJHGFDSAzxcvbnm,./?><MNBVCXZ`~!@#$%^&*()_+=-|\", HeightTB) Then
  MsgBox "Invalid property value : " + HeightTB, vbCritical, "Error"
  HeightTB = R$
 End If
 R$ = HeightTB
End Sub

Private Sub NameTB_Change()
 Static R$
 If StrInText(" /\|()!@#$%^&*-=+[]{}<>?~`';:,.""", NameTB) Then MsgBox "Not a legal object name :  " + NameTB, vbCritical, "Error": NameTB = R$
 R$ = NameTB
End Sub

Private Sub Sensor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 VoteHeld = True
 If VoteHeld = True And X < Sensor.Width + 30 And X > 0 Then
  Rating.Width = X
 End If
End Sub

Private Sub Sensor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If VoteHeld = True And X < Sensor.Width + 30 And X > 0 Then
  Rating.Width = X
 End If
End Sub

Private Sub Sensor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 VoteHeld = False
 MsgBox "Sorry, because of the changes in PSC voting system, it won't work this way, please click on the button to goto PSC and vote/comment for this code, thanks."
End Sub

Private Sub TabContent_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  'Pass the event to the Design Sheet
  DS_KeyDown KeyCode, Shift
End Sub

Private Sub TabMO_Click(Index As Integer)
 TabImg.Picture = TabPic(Index).Picture
 For i = 0 To 2
  TabContent(i).Visible = False
 Next i
 TabContent(Index).Visible = True
End Sub

Private Sub RedrawAll()
  DS.Cls
  For j = 0 To 55
   If Points(j, 0).X = -1 Then GoTo 1
   px = Points(j, 0).X
   py = Points(j, 0).Y
   lx = px: ly = py: hx = 0: hy = 0
   DS.Line (px - 30, py - 30)-(px + 30, py + 30), RGB(0, 255, 0), B
   DS.PSet (px, py)
   For i = 0 To 400
    If Points(j, i).X = -1 Then Exit For
    px = Points(j, i).X
    py = Points(j, i).Y
    If j = SelPG Then
     If px > hx Then hx = px
     If py > hy Then hy = py
     If px < lx Then lx = px
     If py < ly Then ly = py
    End If
    If i = SelP(j) And j = SelPG Then c = RGB(0, 0, 255): Text3.Text = Str$(px \ 15): Text2.Text = Str$(py \ 15) Else c = RGB(0, 255, 0)
    DS.Line -(px, py), RGB(255, 0, 0)
    DS.Line (px - 30, py - 30)-(px + 30, py + 30), c, B
    DS.PSet (px, py)
   Next i
   If j = SelPG Then
    BB.Left = lx
    BB.Top = ly
    BB.Width = hx - lx
    BB.Height = hy - ly
    BB.Visible = True
   End If
1 Next j
 If SelPG = -1 Then Text1.Text = "" Else Text1.Text = Str$(SelPG)
 Label2(8).Caption = OMS(OvlMethod(SelPG))
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Text1.Text = Str$(Val(Text1.Text))
  If Val(Text1.Text) < 0 Or Val(Text1.Text) > 55 Then Text1.Text = "0"
  SP = Val(Text1.Text)
1 If Points(SP, 0).X > -1 Then SelPG = SP Else SP = SP - 1: GoTo 1
  RedrawAll
 End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Points(SelPG, SelP(SelPG)).Y = Val(Text2.Text) * 15
  RedrawAll
 End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Points(SelPG, SelP(SelPG)).X = Val(Text3.Text) * 15
  RedrawAll
 End If
End Sub

Private Sub TextBox1_Change()
 
End Sub

Private Sub Timer1_Timer()
 'Blink the voting bar prompt to attract YOUR attention! :-)
 If Label27.ForeColor = RGB(255, 0, 0) Then Label27.ForeColor = RGB(0, 0, 0) Else Label27.ForeColor = RGB(255, 0, 0)
End Sub

Private Function FileExists(FileName As String) As Boolean
 On Error GoTo 1
 Open FileName For Input As #1
  Input #1, R$
 Close #1
 FileExists = True
 Exit Function

1 FileExists = False
  Close #1
End Function

Private Sub BuildFormFile(FileName As String)
 Open FileName For Output As #1
   If CB2.ListIndex = 1 Then R$ = "0 'None" Else R$ = "1 'Fixed Single"
   Print #1, "VERSION 5.00"
   Print #1, "Begin VB.Form " + NameTB
   Print #1, "  Appearance      =" + Str$(CB1.ListIndex)
   Print #1, "  Autoredraw      =" + Str$(CheckBox1.Value)
   Print #1, "  BorderStyle     = " + R$
   Print #1, "  Caption         = """ + NameTB + """"
   Print #1, "  ClientHeight    = " + HeightTB
   Print #1, "  ClientLeft      = " + LeftTB
   Print #1, "  ClientTop       = " + TopTB
   Print #1, "  ClientWidth     = " + WidthTB
   Print #1, "  ClipControls    =" + Str$(CheckBox2.Value)
   Print #1, "  ControlBox      =" + Str$(CheckBox5.Value)
   Print #1, "  LinkTopic       = """ + NameTB + """"
   Print #1, "  ScaleHeight     = " + HeightTB
   Print #1, "  ScaleWidth      = " + WidthTB
   Print #1, "  ShowInTaskbar   =" + Str$(CheckBox3.Value)
   Print #1, "  StartUpPosition = 3 'Windows Default"
   Print #1, "End"
   Print #1, "Attribute VB_Name = " + NameTB
   Print #1, "Attribute VB_GlobalNameSpace = False"
   Print #1, "Attribute VB_Creatable = False"
   Print #1, "Attribute VB_PredeclaredId = True"
   Print #1, "Attribute VB_Exposed = False"
   Print #1,
   Print #1, "'Paradise Form Vision Build: " + Date$ + " at " + Time$
   Print #1,
   Print #1, "Private Type PointXY"
   Print #1, "  X As Long"
   Print #1, "  Y As Long"
   Print #1, "End Type"
   Print #1,
   Print #1, "Private Const RGN_AND = 1"
   Print #1, "Private Const RGN_COPY = 5"
   Print #1, "Private Const RGN_DIFF = 4"
   Print #1, "Private Const RGN_OR = 2"
   Print #1, "Private Const RGN_XOR = 3"
   Print #1,
   Print #1, "Private Declare Function CreateRectRgn Lib ""gdi32"" Alias ""CreateRectRgn"" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"
   Print #1, "Private Declare Function CombineRgn Lib ""gdi32"" Alias ""CombineRgn"" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long"
   Print #1, "Private Declare Function SetWindowRgn Lib ""user32"" Alias ""SetWindowRgn"" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long"
   Print #1, "Private Declare Function CreatePolygonRgn Lib ""gdi32"" Alias ""CreatePolygonRgn"" (lpPoint As PointXY, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long"
   Print #1,
   Print #1, "Private Sub Form_Activate()"
   If DS.Tag <> "" Then Print #1, "  Me.Picture = LoadPicture(""" + DS.Tag + """)"
   If CB2.ListIndex = 1 Then Print #1, "  Call DrawShape()"
   Print #1, "End Sub"
   If CB2.ListIndex = 1 Then
   Print #1,
   Print #1, "'And here's the precedure which creates the shape of the form"
   Print #1,
   Print #1, "Private Sub DrawShape()"
   Print #1, "  Dim Points(490) As PointXY"
   Print #1, "  hRgn& = CreateRectRgn(0, 0, 0, 0)"
   For j = 0 To 55
    Select Case OvlMethod(j)
     Case 0
      R$ = "RGN_OR"
     Case 1
      R$ = "RGN_COPY"
     Case 2
      R$ = "RGN_AND"
     Case 3
      R$ = "RGN_DIFF"
     Case 4
      R$ = "RGN_XOR"
     Case Else
      'That's impossible !!
    End Select
    If Points(j, 0).X = -1 Then GoTo 3
    For i = 0 To 490
     If Points(j, i).X = -1 Then Exit For
     Print #1, "  Points(" + Right$(Str$(i), Len(Str$(i)) - 1) + ").X =" + Str$(Points(j, i).X / 16)
     Print #1, "  Points(" + Right$(Str$(i), Len(Str$(i)) - 1) + ").Y =" + Str$(Points(j, i).Y / 16)
    Next i
    Print #1, "  hPGRgn& = CreatePolygonRgn(Points(0)," + Str$(i) + ", 1)"
    Print #1, "  CombineRgn hRgn&, hRgn&, hPGRgn&, " + R$
3  Next j
   Print #1, "  SetWindowRgn Me.hWnd, hRgn&, True"
   Print #1, "End sub"
   End If
  Close #1
 'End If
  
 'Close #1
End Sub

Private Function StrInText(Str As String, TextStr As String) As Boolean
 'Check if <TextStr> contains a single character out of <Str>
 temp = False
 For i = 1 To Len(Str)
  If InStr(TextStr, Mid$(Str, i, 1)) > 0 Then temp = True
 Next i
 StrInText = temp
End Function

