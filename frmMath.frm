VERSION 5.00
Begin VB.Form frmMath 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   Caption         =   "Math checker"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "5"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "5"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "5"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "5"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "5"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "15"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "15"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "15"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "15"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "15"
      Top             =   1440
      Width           =   735
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00C0FFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   20
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   19
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00C0FFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   18
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00C0FFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdDivide 
      BackColor       =   &H00C0FFC0&
      Caption         =   "(/)Divide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdMultiply 
      BackColor       =   &H00C0FFC0&
      Caption         =   "(*)Multiply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdMinus 
      BackColor       =   &H00C0FFC0&
      Caption         =   "(-)Minus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlus 
      BackColor       =   &H00C0FFC0&
      Caption         =   "(+) Plus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdNumGenerator 
      BackColor       =   &H0080FF80&
      Caption         =   "Number Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   5160
      TabIndex        =   11
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2880
      TabIndex        =   7
      ToolTipText     =   "You May use Tab key to go down"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2880
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "5"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "15"
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total Question "
      Height          =   375
      Left            =   5880
      TabIndex        =   53
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   52
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label lblCalc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Click > to Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   51
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1920
      TabIndex        =   50
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "What To Do :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3720
      TabIndex        =   38
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3720
      TabIndex        =   37
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   36
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3720
      TabIndex        =   35
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   34
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   33
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   32
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   31
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   30
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   29
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   28
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   27
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   26
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   25
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   24
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2520
      TabIndex        =   23
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   22
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblPercentCorrect 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   21
      Top             =   3960
      Width           =   5295
   End
   Begin VB.Label lblIncorrectScore 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblCorrectScore 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'To calculate and show value in Text and Labels
Dim i As Integer
'To compare correct result with entered values in Text3
Dim J As Integer
Dim intArray(5) As Integer
Dim intCheck(5) As Integer
'
Dim blnPlus As Boolean
Dim blnMinus As Boolean
Dim blnMultiply As Boolean
Dim blndivide As Boolean
'
Dim blnNumgenerator As Boolean
'
Dim optNum As Integer 'Used to show What to do
Dim strOptType As String
'
Dim intCorrectScore As Integer
Dim intInCorrectScore As Integer
Private Sub cmdDivide_Click()
Dim strDivide1(5) As String
Dim strDivide2(5) As String
Static intattempt As Integer
intattempt = intattempt + 1
'List1.AddItem "Attempt " & "--Divide---" & intattempt
For i = 0 To 5
strDivide1(i) = (Val(Text1(i).Text))
intArray(i) = Format((Val(Text1(i).Text) / Val(Text2(i).Text)), "#0.00")
Label1(i).Caption = Format((Val(Text1(i).Text) / Val(Text2(i).Text)), "#0.00")
Verify
Next i
List1.AddItem "Division" & Space$(15) & intattempt & Space$(15) & intCorrectScore & Space$(15) & intInCorrectScore

cmdNumGenerator.Enabled = True
cmdDivide.Visible = False
showNext
intCorrectScore = 0
intInCorrectScore = 0
End Sub
Private Sub cmdMinus_Click()
On Error Resume Next
Static intattempt As Integer
On Error Resume Next
intattempt = intattempt + 1
'List1.AddItem "Attempt " & "--Minus---" & intattempt
For i = 0 To 5
intArray(i) = (Val(Text1(i).Text) - Val(Text2(i).Text))
Label1(i).Caption = intArray(i)
Verify
Next i
If intCorrectScore = 6 Then
MsgBox " Well done you got all correct"
Else
MsgBox "you got  " & intInCorrectScore & "  wrong"
End If
List1.AddItem "Minus    " & Space$(15) & intattempt & Space$(15) & intCorrectScore & Space$(15) & intInCorrectScore
cmdNumGenerator.Enabled = True
cmdMinus.Visible = False
showNext
intCorrectScore = 0
intInCorrectScore = 0
End Sub
Private Sub cmdMultiply_Click()
On Error Resume Next
Static intattempt As Integer
intattempt = intattempt + 1
'List1.AddItem "Attempt " & "--Multiply---" & intattempt
For i = 0 To 5
intArray(i) = (Val(Text1(i).Text) * Val(Text2(i).Text))
Label1(i).Caption = intArray(i)
Verify
Next i
If intCorrectScore = 6 Then
MsgBox " Well done you got all correct"
Else
MsgBox "you got  " & intInCorrectScore & "  wrong"
End If
cmdNumGenerator.Enabled = True
Cls
List1.AddItem "Multiply " & Space$(15) & intattempt & Space$(15) & intCorrectScore & Space$(15) & intInCorrectScore
cmdMultiply.Visible = False
intCorrectScore = 0
intInCorrectScore = 0
showNext
End Sub
Private Sub cmdNumGenerator_Click()
On Error Resume Next
'Number for Text1
Dim intNum1(5) As Integer
'Number for Text2
Dim intNum2(5) As Integer
Dim iPlus As Integer
'let Label3 show type of math
showTypeofMath
If optNum = 4 Then
ClearUp
NumDivide

ElseIf optNum = 3 Then
ClearUp
numMultiply
Else
ClearUp
'
'Text1(iPlus).Text = ""
For iPlus = 0 To 5
'
intNum1(iPlus) = Int((15 * Rnd) + 10)
intNum2(iPlus) = Int((10 * Rnd) + 2)
Text1(iPlus).Text = intNum1(iPlus)
Text2(iPlus).Text = intNum2(iPlus)
Next iPlus
If optNum = 1 Then
cmdPlus.Visible = True
blnPlus = False
Else
cmdPlus.Visible = False
End If
If optNum = 2 Then
cmdMinus.Visible = True
blnMinus = False
Else
cmdMinus.Visible = False
End If
End If
'

cmdNumGenerator.Enabled = False
blnNumgenerator = True
showNext
End Sub

Private Sub Verify()
On Error Resume Next
Static intMarks As Integer
Static intWrong As Integer
Static intTotal As Integer
Dim int50 As Integer
'For J = 0 To 5
intCheck(J) = Val(Text3(i).Text)
'Next J
If intArray(i) = intCheck(J) Then
Text3(i).BackColor = vbGreen
intMarks = intMarks + 1
intCorrectScore = intCorrectScore + 1
lblCorrectScore.Caption = "Correct Score  " & intMarks
Else
Text3(i).BackColor = vbRed
intWrong = intWrong + 1
intInCorrectScore = intInCorrectScore + 1
lblIncorrectScore.Caption = "Incorect Score " & intWrong
End If
intTotal = intWrong + intMarks
int50 = ((intMarks / intTotal) * 100)

If int50 > 50 Then
lblPercentCorrect.Caption = "Percent Correct  " & _
Format((((intMarks) / intTotal) * 100), "00.00") & "   %"
  Else
lblPercentCorrect.Caption = " Sorry, Your score is less than 50 % "
End If

lblTotal.Caption = " Total (In correct + Correct)=  " & intTotal
End Sub

Private Sub cmdPlus_Click()
Static intattempt As Integer
On Error Resume Next
intattempt = intattempt + 1
'List1.AddItem "Attempt " & "--Plus---" & intattempt
For i = 0 To 5
intArray(i) = (Val(Text1(i).Text) + Val(Text2(i).Text))
Label1(i).Caption = intArray(i)
Verify
Next i
intArray(i) = 0
intCheck(J) = 0
If intCorrectScore = 6 Then
MsgBox " Well done you got all correct"
Else
MsgBox "you got " & intInCorrectScore & "wrong"
End If
List1.AddItem "Plus      " & Space$(15) & intattempt & Space$(15) & intCorrectScore & Space$(15) & intInCorrectScore
cmdNumGenerator.Enabled = True
cmdPlus.Visible = False
intCorrectScore = 0
intInCorrectScore = 0
showNext
End Sub

Private Sub Form_Load()
cmdNumGenerator_Click
cmdNumGenerator.Enabled = False
cmdDivide.Visible = False
cmdPlus.Visible = False
cmdMinus.Visible = False
cmdMultiply.Visible = False
'List1.AddItem "Math Type  " & Space$(2) & "Attempt No.  " & Space$(2) & "Correct  " & Space$(2) & "Incorrect"
'List1.AddItem "---------------------------------------------------------------------"
Label5.Caption = "Math Type  " & Space$(2) & "Attempt No.  " & Space$(2) & "Correct  " & Space$(2) & "Incorrect"
End Sub

Private Sub optSelect_Click(Index As Integer)
Refresh
On Error Resume Next
Select Case Index
Case 0 ' Show Plus
 If cmdPlus.Visible = False Then cmdPlus.Visible = True
 If cmdMinus.Visible = True Then cmdMinus.Visible = False
 If cmdDivide.Visible = True Then cmdDivide.Visible = False
 If cmdMultiply.Visible = True Then cmdMultiply.Visible = False
'
optNum = 1
strOptType = "+"
'
Case 1 ' Show Minus
optNum = 2
strOptType = "-"
If cmdMinus.Visible = False Then cmdMinus.Visible = True
'
If cmdMultiply.Visible = True Then cmdMultiply.Visible = False
If cmdPlus.Visible = True Then cmdPlus.Visible = False
If cmdDivide.Visible = True Then cmdDivide.Visible = False

Case 2 ' Show multiply
optNum = 3
strOptType = "*"
If cmdMultiply.Visible = False Then cmdMultiply.Visible = True
'
If cmdPlus.Visible = True Then cmdPlus.Visible = False
If cmdMinus.Visible = True Then cmdMinus.Visible = False
If cmdDivide.Visible = True Then cmdDivide.Visible = False
 '

'
Case 3 ' Show divide
optNum = 4
strOptType = "/"
If cmdDivide.Visible = False Then cmdDivide.Visible = True
'
If cmdPlus.Visible = True Then cmdPlus.Visible = False
If cmdMinus.Visible = True Then cmdMinus.Visible = False
If cmdMultiply.Visible = True Then cmdMultiply.Visible = False

End Select
cmdNumGenerator_Click
blnNumgenerator = False

End Sub
Private Sub NumDivide()
Randomize
Dim s As Integer
Dim intDiv1(5) As Integer
Dim intDiv2(5) As Integer
Dim intDiv3(5) As Integer
Call Cls
For s = 0 To 5
intDiv2(s) = Int((10 * Rnd) + 2)
intDiv1(s) = Int((6 * Rnd) + 2)
intDiv3(s) = intDiv2(s) * intDiv1(s)
Text1(s).Text = intDiv3(s)
Text2(s).Text = intDiv2(s)
Next s
cmdDivide.Visible = True
blndivide = False

End Sub
Private Sub ClearUp()
Dim intClean As Integer
For intClean = 0 To 5
Label1(intClean).Caption = ""
Text3(intClean).Text = ""
Text3(intClean).BackColor = vbYellow
Next intClean
End Sub
Private Sub numMultiply()
Dim sM As Integer
Dim intMul1(5) As Integer
Dim intMul2(5) As Integer
Dim intMul3(5) As Integer
Call Cls
For sM = 0 To 5
intMul2(sM) = Int((8 * Rnd) + 2)
intMul1(sM) = Int((8 * Rnd) + 2)
intMul3(sM) = intMul2(sM) * intMul1(sM)
Text1(sM).Text = intMul1(sM)
Text2(sM).Text = intMul2(sM)
Next sM

cmdMultiply.Visible = True
blnMultiply = False

End Sub
Private Sub showTypeofMath()
Dim intShowMath As Integer
For intShowMath = 0 To 5
Label2(intShowMath).Caption = strOptType
Next
End Sub
Private Sub showNext()
Dim stroptTypeClone As String
If blnNumgenerator = True Then
lblMessage.Caption = "Select Math Type"
   
Else
lblMessage.Caption = "Click on Number Generator"
End If
End Sub

Private Sub Text3_KeyDown(IndexDown As Integer, KeyCode As Integer, Indexup As Integer)
On Error Resume Next
IndexDown = IndexDown + 1
If KeyCode = vbKeyDown Then
 Text3(IndexDown).SetFocus
End If
'
'next

End Sub

Private Sub Text3_KeyUp(Indexup As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
'Indexup = Text3(Indexup)
Indexup = Indexup - 1
If KeyCode = vbKeyUp Then
 Text3(Indexup).SetFocus
End If
End Sub
