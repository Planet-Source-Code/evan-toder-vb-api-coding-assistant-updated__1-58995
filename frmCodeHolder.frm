VERSION 5.00
Begin VB.Form frmCodeHolder 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox graphicsTextToScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Text            =   "frmCodeHolder.frx":0000
      Top             =   4995
      Width           =   6765
   End
   Begin VB.TextBox graphicsColorLongToRgb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Text            =   "frmCodeHolder.frx":155E
      Top             =   4725
      Width           =   6765
   End
   Begin VB.TextBox generalIniFileManipulationClass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Text            =   "frmCodeHolder.frx":1749
      Top             =   4725
      Width           =   6765
   End
   Begin VB.TextBox generalTextboxNumbersOnly 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Text            =   "frmCodeHolder.frx":1EE2
      Top             =   4230
      Width           =   6765
   End
   Begin VB.TextBox generalTwipToPix 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "frmCodeHolder.frx":1F3F
      Top             =   3870
      Width           =   6765
   End
   Begin VB.TextBox windowFormIsLoaded 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "frmCodeHolder.frx":200C
      Top             =   3375
      Width           =   6765
   End
   Begin VB.TextBox graphicsColorToHtml 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "frmCodeHolder.frx":2195
      Top             =   2970
      Width           =   6765
   End
   Begin VB.TextBox subclassBasic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmCodeHolder.frx":24D6
      Top             =   2520
      Width           =   6765
   End
   Begin VB.TextBox arrayIsInitialized 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "frmCodeHolder.frx":2952
      Top             =   2115
      Width           =   6765
   End
   Begin VB.TextBox windowMoveWithoutTitltebar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmCodeHolder.frx":2AC1
      Top             =   1620
      Width           =   6765
   End
   Begin VB.TextBox browserGetHwnd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmCodeHolder.frx":2C5F
      Top             =   1080
      Width           =   6765
   End
   Begin VB.TextBox menuMakeColumns 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmCodeHolder.frx":30E8
      Top             =   765
      Width           =   6765
   End
   Begin VB.TextBox windowZorder 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmCodeHolder.frx":379D
      Top             =   315
      Width           =   6585
   End
End
Attribute VB_Name = "frmCodeHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'this forms purpose is merely to act as a container for
'functions
'each individual function is placed in a seperate textbox
'when the user clicks on the men item "window z-order"
'(under api function menu) then the code that is in the proper
'textbox is copied to the clipboard
Private Sub Form_Load()

End Sub
