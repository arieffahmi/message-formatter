VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Frame fraStyle 
      Caption         =   "Style"
      Height          =   2415
      Left            =   8040
      TabIndex        =   9
      Top             =   360
      Width           =   1695
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italic"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Underline"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.OptionButton optBlack 
      Caption         =   "Black"
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.OptionButton optGreen 
      Caption         =   "Green"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton optBlue 
      Caption         =   "Blue"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame fraColor 
      Caption         =   "Color"
      Height          =   2415
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   1575
      Begin VB.OptionButton optRed 
         Caption         =   "Red"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtMessage 
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   5655
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click Me"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   4920
      Width           =   855
   End
   Begin VB.Image imgBig 
      Height          =   1695
      Left            =   240
      Picture         =   "Message Formatter.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Click here"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Image imgLittle 
      Height          =   1335
      Left            =   480
      Picture         =   "Message Formatter.frx":68F8B
      Stretch         =   -1  'True
      ToolTipText     =   "Click here"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblMessage 
      Height          =   2055
      Left            =   3240
      TabIndex        =   17
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "Message:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Your name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkBold_Click()
    'Change the message text to/from bold
    
    lblMessage.Font.Bold = chkBold.Value
End Sub

Private Sub chkItalic_Click()
    'Change the message text to/from italic
    
    lblMessage.Font.Italic = chkItalic.Value
End Sub

Private Sub chkUnderline_Click()
    'Change the message text to/from underline
    
    lblMessage.Font.Underline = chkUnderline.Value
End Sub

Private Sub cmdClear_Click()
    'Clear the text controls
    
    With txtName
        .Text = ""          'Clear the text box
        .SetFocus           'Reset the insertion point
    End With
    
    lblMessage.Caption = ""
    txtMessage.Text = ""
End Sub

Private Sub cmdDisplay_Click()
    'Display the text in the message area
    
    lblMessage.Caption = txtName.Text & ": " & txtMessage.Text
End Sub

Private Sub cmdExit_Click()
    'Exit the project
    
    End
End Sub

Private Sub cmdPrint_Click()
    'Print the form
    
    PrintForm
End Sub

Private Sub imgBig_Click()
    'Switch the icon
    
    imgBig.Visible = False
    imgLittle.Visible = True
End Sub

Private Sub imgLittle_Click()
    'Switch the icon
    
    imgBig.Visible = True
    imgLittle.Visible = False
End Sub
