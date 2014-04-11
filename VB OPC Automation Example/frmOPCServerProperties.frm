VERSION 5.00
Begin VB.Form frmOPCServerProperties 
   Caption         =   "OPC Server Properties"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox BuildNumber 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox MinorVersion 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox MajorVersion 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox ServerState 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox LastUpdate 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox CurrentTime 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox StartTime 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Vendor 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox ProgID 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label9 
      Caption         =   "Build Number:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Minor Version:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Major Version:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Server State:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Last Update Time:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Current Time:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Start Time:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Vendor:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Prog ID:"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmOPCServerProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The OPC Server Properties form is a simple display
' only form.  Most of the properties found on the
' actual OPCServer object of the Automation wrapper
' are read only.  This form accesses these properties
' via OPCServerClass object supplied in this example.
'
Option Explicit
Option Base 1


' Update the form controls using data from the
' OPCServerClass object.
'
Private Sub Form_Load()
    Dim VarString As String
    Dim VarTime As Date
    Dim VarInt As Integer
    Dim VarLong As Long
    
    ' This form should be able to be displayed
    ' without the Module1.SelectedOPCServer being
    ' set but it's always a good idea to test it.
    If Not Module1.SelectedOPCServer Is Nothing Then
        ' Load Server Name This is the Prog ID of
        ' the server.
        If (Module1.SelectedOPCServer.GetServerName(VarString) = True) Then
            ProgID.Text = VarString
        Else
            ProgID.Text = "No ID"
        End If
    
        ' The Vendor info will usually be either the
        ' primary manufacturer of the OPC server or
        ' possibly the vendor name of the OPC toolkit
        ' used by the OPC server vendor/developer.
        If (Module1.SelectedOPCServer.GetVendorInfo(VarString) = True) Then
            Vendor.Text = VarString
        Else
            Vendor.Text = "No Vendor"
        End If
        
        ' The server Start Time is the time this OPC
        ' server as start on the host machine.  This
        ' time will be the same for all connected
        ' OPC clients.
        If (Module1.SelectedOPCServer.GetStartTime(VarTime) = True) Then
            StartTime.Text = VarTime
        Else
            StartTime.Text = ""
        End If
        
        ' The current time reports the time at the OPC
        ' servers host machine.  This time is not
        ' guaranteed to be the same for OPC clients
        If (Module1.SelectedOPCServer.GetCurrentTime(VarTime) = True) Then
            CurrentTime.Text = VarTime
        Else
            CurrentTime.Text = ""
        End If
        
        ' The Last Update time reports the time
        ' last data update occured for this instance
        ' of the OPCServer object.
        If (Module1.SelectedOPCServer.GetLastUpdateTime(VarTime) = True) Then
            LastUpdate.Text = VarTime
        Else
            LastUpdate.Text = ""
        End If
        
        ' The Server State returns the operating condition
        ' of the OPC Server.  There are currently 6 states
        ' an OPC server can be in.
        ' OPC_STATUS_RUNNING = 1 (Server running normally)
        ' OPC_STATUS_FAILED  = 2 (Vendor specific failure has
        '                         occured, all intefaces should
        '                         return E_FAIL.)
        ' OPC_STATUS_NOCONFIG =3 (Server running but no
        '                         configuration info available)
        ' OPC_STATUS_SUSPENDED=4 (Server is suspended and is not
        '                         reading or writing data, data
        '                         will returned as Out of Service)
        ' OPC_STATUS_TEST    = 5 (Server in test mode, outputs
        '                         disabled)
        ' OPC_STATUS_DISCONNECTED = 6 (Server has disconnected)
        If (Module1.SelectedOPCServer.GetServerState(VarLong) = True) Then
            Select Case VarLong
                Case 1
                    ServerState.Text = "Server Running"
                Case 2
                    ServerState.Text = "Server Failed"
                Case 3
                    ServerState.Text = "Server No Configuraition"
                Case 4
                    ServerState.Text = "Server Suspended"
                Case 5
                    ServerState.Text = "Server Test Mode"
                Case 6
                    ServerState.Text = "Server Disconnected"
            End Select
        Else
            ServerState.Text = ""
        End If
    
        ' The next set of calls simple get the version number
        ' of the connected OPC server.  This can be useful if
        ' your application is using a specific OPC server that
        ' may revision changes you need to know about.
        If (Module1.SelectedOPCServer.GetMajorVersion(VarInt) = True) Then
            MajorVersion.Text = VarInt
        Else
            MajorVersion.Text = ""
        End If
        
        If (Module1.SelectedOPCServer.GetMinorVersion(VarInt) = True) Then
            MinorVersion.Text = VarInt
        Else
            MinorVersion.Text = ""
        End If
        
        If (Module1.SelectedOPCServer.GetBuildNumber(VarInt) = True) Then
            BuildNumber.Text = VarInt
        Else
            BuildNumber.Text = ""
        End If
    End If
       
      
End Sub


' Since there nothing set the OK button simple
' unloads the form.
'
Private Sub OkButton_Click()
    ' Reenable the tree view on the main form
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
    frmOPCServerProperties.Show
End Sub

