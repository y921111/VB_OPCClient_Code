VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   3840
   End
   Begin VB.Frame fraMainFrame 
      Height          =   3630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7380
      Begin VB.PictureBox picLogo 
         Height          =   1545
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   1485
         ScaleWidth      =   7155
         TabIndex        =   1
         Top             =   1080
         Width           =   7215
      End
      Begin VB.Label Purpose 
         AutoSize        =   -1  'True
         Caption         =   "Visual Basic OPC Example"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   5
         Tag             =   "Product"
         Top             =   600
         Width           =   3045
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "KEPware, Inc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4200
         TabIndex        =   4
         Tag             =   "CompanyProduct"
         Top             =   120
         Width           =   2385
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   3000
         Width           =   1875
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright  2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   2
         Tag             =   "Copyright"
         Top             =   3240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Here's that sometimes annoying splash screen code.  The
' useful tidbits here are the use of the Windows API
' function SetWindowPos to keep the window on top while
' it is displayed and the use of a timer to shut is down
' after a brief period.

Option Explicit

' Declare a function pointer to the USER32 lib function
' SetWindowPos.  To get the actual parameters for SetWindowPos
' see your windows api help file.
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Load()
    ' Build the application version string and display it
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    ' The windows API call SetWindowPos is used to keep the
    ' splash screen form on top by setting it to be the top
    ' most window. One thing to be aware of when using this
    ' method to keep a form on the top is that it can
    ' potentially hide other forms or popups like the
    ' message box behind a form once this call is made
    ' so use it with care.  This can have the effect of
    ' your application appearing to be locked up.
    
    SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
End Sub
        
' Simple one shot handling of the timer event to remove
' the form from the display. Currently this is set for 4
' seconds.
'
Private Sub Timer1_Timer()
    Unload frmSplash
End Sub

' frmSplash display also will be removed when you click it.
'
Private Sub fraMainFrame_Click()
    Unload frmSplash
End Sub

