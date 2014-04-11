VERSION 5.00
Begin VB.Form frmItemUpdateInterval 
   Caption         =   "Item Display Update Interval"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Hit Ok to accept the new update interval"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Interval 
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "50"
      ToolTipText     =   "Enter the update interval for the display of data (50 - 10000) ms."
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Interval (ms):"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmItemUpdateInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form is used to change the display update
' interval of the timer object used to update the
' lvListView of the frmMain form.  Using a timer
' to update the data displayed in the list view
' control is a fairly low tech way to reduce the
' flicker on this control.  This is by no means
' perect. There is a way to make the display very
' smooth but it involves sub classing the ListView
' control and goes well beyond what I want to present
' in this demo.

Option Explicit
Option Base 1


' Get the current time interval of the timer object
' of the frmMain form.  This is done through the
' pointer to the instance of the frmMain object.
'
Private Sub Form_Load()
    Interval.Text = Mid(Str(fMainForm.Timer1.Interval), 2)
End Sub


' Once a new value has been set check the range and
' clip it accordingly before applying it to the
' timer object on frmMain.
'
Private Sub Ok_Click()
    Dim IntVal As Long
    IntVal = Val(Interval.Text)
    If IntVal < 50 Then
        IntVal = 50
    ElseIf IntVal > 10000 Then
        IntVal = 10000
    End If
    Interval.Text = Mid(Str(IntVal), 2)
    
    ' In order to make sure the new interval value takes
    ' effect immediately you need to disable the timer
    ' set the value then reenable it.
    fMainForm.Timer1.Enabled = False
    fMainForm.Timer1.Interval = IntVal
    fMainForm.Timer1.Enabled = True
Unload Me
End Sub


' The Deactivate event is used here as a cheap way
' of keeping the form on top of the application if
' the user clicks on the main window.  There may be
' other ways to do this but this works fairly well.
'
Private Sub Form_Deactivate()
    ' use the show memthod to force the form
    ' back to the front
    frmItemUpdateInterval.Show
End Sub

