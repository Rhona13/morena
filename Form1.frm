VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   12120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12120
   ScaleWidth      =   19890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xretBalance 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14880
      TabIndex        =   43
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox xremarks 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   42
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox xtBalance 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   40
      Top             =   7920
      Width           =   2655
   End
   Begin VB.TextBox xname 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   36
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox xlname 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   35
      Top             =   9240
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   34
      Top             =   11040
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "COMPUTE"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   33
      Top             =   10440
      Width           =   2535
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14880
      TabIndex        =   31
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox xshowRemark 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14880
      TabIndex        =   29
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox xreturn 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14880
      TabIndex        =   27
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   25
      Top             =   9000
      Width           =   2655
   End
   Begin VB.TextBox xbBalance 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   23
      Top             =   9600
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14880
      TabIndex        =   21
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   19
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox onHand 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   17
      Top             =   6480
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox borrow 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   12
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   10
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   8
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000007&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17520
      TabIndex        =   4
      Top             =   11040
      Width           =   1575
   End
   Begin VB.TextBox xlastName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox xfname 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label xreturnBalance 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Return Balance"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12000
      TabIndex        =   44
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Total Balance"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   600
      TabIndex        =   41
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Firstname"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   720
      TabIndex        =   39
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Lastname"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   600
      TabIndex        =   38
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Borrower Details"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   480
      TabIndex        =   37
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Date Returned"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12720
      TabIndex        =   32
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12720
      TabIndex        =   30
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Qty Returned"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12720
      TabIndex        =   28
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Date Borrowed"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6120
      TabIndex        =   26
      Top             =   8880
      Width           =   2655
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Borrowed Balance"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6120
      TabIndex        =   24
      Top             =   9600
      Width           =   2895
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Property Code"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12360
      TabIndex        =   22
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   20
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Quantity On Hand"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5760
      TabIndex        =   18
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Unit Price"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Qty Borrowed"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Product Code"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   9
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Transaction ID"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Name Of The Borrower"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Lastname"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Firstname"
      BeginProperty Font 
         Name            =   "@Malgun Gothic Semilight"
         Size            =   15
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()

xname = xfname.Text
xlname = xlastName.Text
xshowRemark = xremarks.Text
xbBalance = Val(borrow.Text) - Val(onHand.Text)
xretBalance = Val(xreturn.Text) + Val(onHand.Text)
xtBalance = Val(xretBalance.Text) + Val(xbBalance.Text)

End Sub

Private Sub Command2_Click()

xfname.Text = ""
xlname.Text = ""
xname.Text = ""
xlastName.Text = ""
xtBalance.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
borrow.Text = ""
Text5.Text = ""
onHand.Text = ""
xremarks.Text = ""
Text8.Text = ""
Text9.Text = ""
Text11.Text = ""
xbBalance.Text = ""
xreturn.Text = ""
xshowRemark.Text = ""
Text14.Text = ""
xretBalance.Text = ""


End Sub

