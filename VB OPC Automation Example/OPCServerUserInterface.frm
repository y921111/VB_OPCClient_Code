VERSION 5.00
Begin VB.Form frmSelectOPCServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton AbortConnection 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Click here to abort making a selection at this time"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton ConnectToOPCServer 
      Caption         =   "Connect to Server"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Click here to attempt a connection to the selected OPC server"
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox SelectedOPCServer 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "No Server Selected"
      ToolTipText     =   "Select an OPC server from the list above or enter a Prog ID here"
      Top             =   3240
      Width           =   5535
   End
   Begin VB.ListBox OPCServerList 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click on one of the available OPC servers listed here"
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmSelectOPCServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The Select OPC Server form allows the you to select
' an OPC server from a list of servers available on your
' machine.  The GetOPCServerList function of the
' OPCServerClass wraps the GetOPCServers function of the
' OPCServer object.  The GetOPCServerList function can take
' an optional Node name to allow access of nodes on remote
' machine as well.  This demo uses servers on the local
' machine.  The Form_Load event creates a temporary
' OPCServerClass object to get the list of OPC Server

Option Explicit
Option Base 1


Private Sub Form_Load()
    ' Create a temporary OPCServerClass object to get the available server list.
    Dim OPCServer As OPCServerClass
    Set OPCServer = New OPCServerClass

    ' Provide a variant list to hold the available OPC Servers
    Dim ServerList As Variant

    ' Get the list if any
    OPCServer.GetOPCServerList ServerList, "192.168.57.140"

    ' Exit if list is null
'    If ServerList = Empty Then
'       GoTo Exit_Form_Load
'    End If

    ' Clear the list control used to display them
    OPCServerList.Clear
    ' Load the list returned into the List box for user selection
    Dim i As Integer
    For i = LBound(ServerList) To UBound(ServerList)
        OPCServerList.AddItem ServerList(i)
    Next i
    
Exit_Form_Load:
      
    ' Delete the temporary OPCServerClass object
    Set OPCServer = Nothing
End Sub


Private Sub AbortConnection_Click()
    ' Reenable the tree view
    fMainForm.tvTreeView.Enabled = True
    Unload Me
End Sub


' The ConnectToOPCServer click event calls the
' AddSelectedOPCServer function of the main form to add
' a server to the application.  The AddSelectedOPCServer
' function is housed in the main form since this is where
' the new server will be added to the tvTreeView object
' of that form.  The AddSelectedOPCServer calls the
' ConnectOPCServer function of the OPCServerClass object.
' When the user selects an OPC server from the list
' the AddSelectedOPCServer function first creates a new
' OPCServerClass object then attempts to connnect to the
' selected OPC server using this new OPCServerClasss object.
' If the addition is successful then the new server will be
' added to the OPCServers collection that is global to the
' application.
'
Private Sub ConnectToOPCServer_Click()
    ' Check and make sure the user attempts to select an
    ' OPC server.
    If (SelectedOPCServer.Text = "No Server Selected") Then
        MsgBox "Please select an OPC Server from the list or enter a Prog ID."
    Else
        ' Call the AddSelectedOPCServer function on the main
        ' form to add this new server by passing the name of the
        ' server.
        fMainForm.AddSelectedOPCServerMain SelectedOPCServer.Text
        ' Reenable the tree view
        fMainForm.tvTreeView.Enabled = True
        Unload Me
    End If
End Sub


' The Deactivate event is used here as a cheap way
' of keeping the form on top of the application if
' the user clicks on the main window.  There may be
' other ways to do this but this works fairly well.
'
Private Sub Form_Deactivate()
    ' use the show memthod to force the form
    ' back to the front
    frmSelectOPCServer.Show
End Sub


' Each time the user clicks on an OPC server update the
' SelectedOPCServer text box.
'
Private Sub OPCServerList_Click()
    SelectedOPCServer.Text = OPCServerList.List(OPCServerList.ListIndex)
End Sub

