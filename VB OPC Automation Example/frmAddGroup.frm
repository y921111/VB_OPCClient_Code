VERSION 5.00
Begin VB.Form frmAddGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add OPC Group"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox GroupName 
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   8
      ToolTipText     =   "Enter a group name here or leave blank to allow the server to provide a group name"
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton AbortAdd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton AddGroup 
      Caption         =   "Add Group"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox ActiveState 
      Alignment       =   1  'Right Justify
      Caption         =   "Active State:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Click here to change the active state of the group"
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox DeadBand 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Enter a new Group DeadBand 0 - 100"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox UpdateRate 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "100"
      ToolTipText     =   "Enter a new Group Update Rate 0 - 2147483647 milliseconds"
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Group Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "DeadBand:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Update Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The Add OPC Group form allows the user to set
' the group name, the update rate, deadband, and
' active state of an OPC group to be added to
' the selected OPC server connection.  When this
' form is displayed the tvTreeView on the main form
' is disable to prevent the user from invoking a
' second Add group or Add Server menu.  This form
' will only be invoke while the user has an existing
' OPC server connection selected.  This means that
' global variable SelectedOPCServer will be set.
'
' One thing you'll notice about the Add OPC Group
' form is that the GroupName textbox control is
' initialized to a blank string.  Normally you can
' name an OPC group yourself but you also have the
' option of allowing the OPC Server to provide you
' with a unique group name automatically.  To do
' this simply leave the GroupName textbox empty.
' This blank group name is then passed to the OPC
' server which will recognize this as a automatic
' group name.  The server will return the name it
' generates as part of the Automation Interface's
' OPCGroup object.  I will explain this further
' over in the OPCServerClass module.  I will get the
' new automatic group name and display it after the
' new group has been added.
'
' The Active state allows you to set whether or not
' the OPC group you intend to add will be actively
' acquiring data or idle.  The active state of an
' OPC group is an important tool in controlling how
' an OPC server is talking to you phsyical
' device.  The amount of data that an OPC server is
' gathering from a device depends on two factors,
' the number of actual OPC items added to the OPC
' groups and active or inactive state of teh group.
' This is an important consideration.  When
' designing your OPC client application you can
' control the operation of the OPC server by adding
' OPC groups and deleting them as need.  This works
' but it forces the OPC server to do a great deal of
' memory allocation and releasing as well as
' potentially interupting the order of how data is
' being polled from the server.  A better way
' to control how much data is being acquired by the
' OPC server is to use the active state of the group.
' By using the active state of the group you can
' essential turn an entire group of items off. When
' this occurs the OPC server will stop scanning
' all items in this group from the physical device.
' This of course will reduce the bandwidth
' requirements of the OPC server.  By using the
' active state of the group instead of adding and
' deleting the groups you allow the OPC server to
' operate at peak performance.
'
' The update rate of a group is a powerful tool when
' using OPC servers.  By using multiple groups to
' segregate your data you can optimize the amount of
' data the OPC server is acquiring.  By placing data
' with differing update requirements in thier
' own groups you can have one group of items that is
' read at a high rate of speed while another group
' can be set to update infrequently.  Using the
' update rate of the group can be a powerful tool
' for your application and since the update rate
' can be changed at any time you can even design
' your application to change the rate based on what
' the user is doing or looking at.

' When you click on the Add Group button on this
' form a function housed in frmMain called
' AddOPCGroup is called with the items entered here
' as parameters.  This function is housed in the
' main form because this is where the the resulting
' group will be added to the tree view.  This
' AddOPCGorup function found in the main form
' demonstrates how to use the actual add group
' function of the OPCServerClass object.

Option Explicit
Option Base 1


' Handles the Cancel button on the Add OPC group
' form.
'
Private Sub AbortAdd_Click()
    ' Reenable the tree view on the main form
    fMainForm.tvTreeView.Enabled = True
    
    Unload Me ' Remove the Add OPC Form
End Sub


' The user wishes to add the group so call the
' AddOPCGroup function in the main form
'
Private Sub AddGroup_Click()
    ' Attempt to add this new group to the
    ' selected OPC server
    fMainForm.AddOPCGroupMain GroupName.Text, (UpdateRate.Text), Val(DeadBand.Text), ActiveState.Value
    
    ' Reenable the tree view
    fMainForm.tvTreeView.Enabled = True
    
    Unload Me ' Remove the Add OPC Form
End Sub


' The Deactivate event is used here as a cheap way
' of keeping the form on top of the application if
' the user clicks on the main window.  There may be
' other ways to do this but this works fairly well.
'
Private Sub Form_Deactivate()
    ' use the show memthod to force the form
    ' back to the front
    frmAddGroup.Show
End Sub
