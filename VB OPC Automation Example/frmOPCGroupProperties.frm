VERSION 5.00
Begin VB.Form frmOPCGroupProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OPC Group Properties"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox ActiveState 
      Alignment       =   1  'Right Justify
      Caption         =   "Active State:"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Click here to change the active state of the group"
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox DeadBand 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "Enter a new Group DeadBand 0 - 100"
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox UpdateRate 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "100"
      ToolTipText     =   "Enter a new Group Update Rate 0 - 2147483647 milliseconds"
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox GroupName 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Group Name"
      ToolTipText     =   "The Group Name cannot be edited in this demo"
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Dead Band:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Update Rate:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Group Name:"
      Height          =   225
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frmOPCGroupProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The OPC Group Properties form allows the user to modify the
' the Update Interval, DeadBand and Active state of a group
' that has already been added to a server.  This form
' demonstrates how these parameters of the OPCGroupClass
' can be directly changed to control how an OPC group is
' operating.  As changes are made on this form they will
' be applied to the group immediately.  The variable
' SelectedOPCGroup is set when the user either Left or Right
' clicks on a group name in the tvTreeView control.  The
' SelectedOPCGroup pointer determines what instance of
' the OPCGroupClass object will be effected.
'
' To see the effects of these changes add some items to the
' OPC group like "Device_1.R0" and "Device_1.R1" then
' make changes to the Update rate and active state to see
' their effect.

Option Explicit
Option Base 1


' Get the current values of the group from the selected
' instance of the OPCGroupClass object.
'
Private Sub Form_Load()
    Dim VarSingle As Single
    Dim VarLong As Long
    Dim VarState As Boolean
        
    If Not Module1.SelectedOPCGroup Is Nothing Then
        ' Get the groups name.  Remeber this could be a name that
        ' has been provided by the OPC server if you added the group
        ' with a blank group name.
        GroupName.Text = Module1.SelectedOPCGroup.GetGroupName
               
        ' Get the Update rate
        If (Module1.SelectedOPCGroup.GetGroupUpdateRate(VarLong) = True) Then
            UpdateRate.Text = Str(VarLong)
        Else
            UpdateRate.Text = "0"
        End If
        
        ' Get the DeadBand
        If (Module1.SelectedOPCGroup.GetGroupDeadBand(VarSingle) = True) Then
            DeadBand.Text = Str(VarSingle)
        Else
            DeadBand.Text = "0"
        End If
        
        ' Get the Active State
        If (Module1.SelectedOPCGroup.GetGroupActiveState(VarState) = True) Then
            ' Convert from the returned Boolean to a what the check box wants.
            If VarState = True Then
                ActiveState.Value = 1
            Else
                ActiveState.Value = 0
            End If
        Else
            ActiveState.Value = 0
        End If
        
    End If
       
End Sub


' Change the Active state of the group by using the
' SetGroupActiveState function of the OPCGroupClass object.
' In your application this may be done programmatically rather
' than being driven from user input. The change takes effect
' immediately.
'
Private Sub ActiveState_Click()
    If Not Module1.SelectedOPCGroup Is Nothing Then
        Module1.SelectedOPCGroup.SetGroupActiveState (ActiveState.Value)
    End If
End Sub


' Change the Dead Band of the group by using the
' SetGroupDeadBand function of the OPCGroupClass object.
' In your application this may be done programmatically rather
' than being driven from user input. The change takes effect
' immediately.
'
Private Sub DeadBand_Change()
    If Not Module1.SelectedOPCGroup Is Nothing Then
        Module1.SelectedOPCGroup.SetGroupDeadBand (Val(DeadBand.Text))
    End If
End Sub


' Change the Update Rate of the group by using the
' SetGroupUpdateRate function of the OPCGroupClass object.
' In your application this may be done programmatically rather
' than being driven from user input. The change takes effect
' immediately. The update rate of a group is a powerful tool
' when using OPC servers.  By using multiple groups to
' segregate your data you can optimize the amount of data
' the OPC server is acquiring.  By placing data with
' differing update requirements in their own groups you can
' have one group of items that is read at a high rate of
' speed while another group can be set to update infrequently.
' Using the update rate of the group can be a powerful tool
' for your application and since the update rate can be changed
' at any time you can even design your application to change
' the rate based on what the user is doing or looking at.
'
Private Sub UpdateRate_Change()
    If Not Module1.SelectedOPCGroup Is Nothing Then
        Module1.SelectedOPCGroup.SetGroupUpdateRate (Val(UpdateRate.Text))
    End If
End Sub


' Since the change on this form take effect immediately there's
' nothing to do on exit.
'
Private Sub OkButton_Click()
    ' Reenable the tree view
    fMainForm.tvTreeView.Enabled = True
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
    frmOPCGroupProperties.Show
End Sub
