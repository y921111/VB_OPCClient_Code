VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "KEPServerEX_OPC_Example"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5880
      Top             =   480
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin ComctlLib.TreeView tvTreeView 
      Height          =   5220
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Right Click Here to add OPC Servers and Groups"
      Top             =   360
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   9208
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.ListView lvListView 
      Height          =   5250
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "With a group selected in the tree view to the left, use the right mouse button to add an item to the group"
      Top             =   360
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   9260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item List:"
         Height          =   270
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Tag             =   "ItemList:"
         Top             =   0
         Width           =   3210
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Server && Groups:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   " Server & Groups:"
         Top             =   0
         Width           =   2010
      End
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5595
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5398
            Text            =   "Status"
            TextSave        =   "Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2006-1-12"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "17:32"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5880
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   6000
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":109A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2134
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewUpdateDispaly 
         Caption         =   "Item Display &Update Interval..."
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About KEPServerEX_OPC_Example..."
      End
   End
   Begin VB.Menu mnuTreeView 
      Caption         =   "TreeViewContext"
      Visible         =   0   'False
      Begin VB.Menu mnuTreeViewNewServer 
         Caption         =   "New &Server Connection..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTreeViewNewGroup 
         Caption         =   "New &Group..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTreeViewNewItem 
         Caption         =   "New &Item"
      End
      Begin VB.Menu mnuTreeContextBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeViewProperties 
         Caption         =   "&Properties..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTreeContextBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeViewDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuListView 
      Caption         =   "ListViewContext"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuListViewNewItem 
         Caption         =   "New &Item..."
      End
      Begin VB.Menu mnuListViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewSetActive 
         Caption         =   "Set &Active"
      End
      Begin VB.Menu mnuListViewSetInactive 
         Caption         =   "Set I&nactive"
      End
      Begin VB.Menu mnuListViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewWriteValue 
         Caption         =   "&Write Value..."
      End
      Begin VB.Menu mnuListViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The main form frmMain is the central user interface focal point
' for this example application. While all of the actual OPC
' interface code is contained in the three class modules
' OPCServerClass, OPCGroupClass, and OPCItemClass, the use
' of these classes is demostrated primarily in this module.
' Think of this module as being the basis for your own OPC
' application.  While your application may not have the same
' user interface requirements, the mechanisms for adding server
' connections, groups, and items is the same.  One thing to
' keep in mind is that while this example applicaiton
' demonstrates the various OPC elements driven by User input
' your application can do all of these steps programtically,
' or via configuration files, or database contents.
'
' The primary bodies of code in this form deal with managing
' user interaction with the tvTreeView and lvListView controls
' on this form.  That management includes maintaining the
' OPC server, group, and item objects stored in these views.
' As you add server connections, groups and items, both the
' class modules and the TreeView and ListView controls are kept in
' sync with these item.  Simply put the TreeView and ListView
' contain a collection of objects but the OPC class modules
' keep their own list so you don't need to have any user interface
' elements for your own OPC application.
'
' As part of managing the user interface, the frmMain module
' also handles determining what OPC server, group or item has been
' selected by the user.  Once determined the three variables in
' Module1.bas,  SelectedOPCServer, SelectedOPCGroup,
' and SelectedOPCItem will be set accordingly.  These variables
' are important in other parts of this example application.  While
' not used in the three class modules, these variables are used in
' the other VB forms of this application to determine what object
' the user has selected for modification or use.
'
' One important element of how the user interface interface for
' this example is the use of the key string used on the Node
' items for the tvTreeView and on the ListItems for the lvListView.
' These key have been given very specific names with a number. In the
' TreeView there are two types of keys by name "Server X" and
' "Group X Y".  For the server key the Key always contains the word
' "Server" plus some number.  For group items the Key always
' contains the word "Group" plus a group number and the server
' number.  For this example it is crucial that these not be changed
' as all object selection is based on using these key to identify
' selected  object.  I give more detail below but it worth mentioning
' the now due to its importance.
'
' Only this form contains a menu bar.  If you have already run the
' application you have seen that there is very little in the menus.
' If you look at the menu bar in the menu editor you will see that
' there a great deal of hidden(popup) menus.  These popup menus are
' invoked by right clicking in either the tvTreeView or the lvListView.
' These right click context menus are used as the primary method of user
' interaction within the application.  Normally you would mimic these
' menus on the menu bar and I would have as well but you then must
' also take care of enabling and disabling them as needed.  I felt that
' the extra code would have clouded the application so I provide
' only the context menus.
'
' At the end of this module is some code generated by the VB app
' wizard, I have placed this code at the end of this module.
' It deals with the resizing of the this applications client
' space but has nothing to do with OPC connectivity.


' By convention, the OPC Automation interface assumes that arrays
' are 1 based. By the way, all vars in this program have to be
' declared explicitly
'
Option Explicit
Option Base 1

' These variables are used in the standard code generated by
' VB to handle moving the splitter bar on the this form.
Dim mbMoving As Boolean
Const sglSplitLimit = 500
Dim LastTopItem As Integer


' On form load we need to get the last size of the application we
' also need to add the heading to the lvListView control.
'
Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 7500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    ' Add the three OPC heading to the lvListView control
    ' The ItemID is the fully qualified Item identification string
    ' for an OPC item.  The Value heading is the current value
    ' of the item. The Status heading is Quality value of the
    ' OPC item.
    lvListView.ColumnHeaders.Add , , "ItemID", lvListView.Width / 3
    lvListView.ColumnHeaders.Add , , "Value", lvListView.Width / 3
    lvListView.ColumnHeaders.Add , , "Status", lvListView.Width / 3
    lvListView.View = lvwReport ' Place list view text report mode
    
    ' This variable is used to determine when the user is attempting
    ' to scroll the OPC Item window.  When the window is scrolled the
    ' currently selected OPC Item must be deselected to allow the
    ' scroll beyond the selected item.  If the selection isn't cleared
    ' the ListView control won't allow the list to be scrolled.  Normally
    ' you don't notice this because the list view isn't normally being
    ' updated continuously.
    LastTopItem = -1
End Sub


' The unload forms sub handles cleaning up for the application.
' From an OPC standpoint it releases each OPCServerClass object
' contained in the OPCServers collection.  Due to the design of
' the OPCServerClass object it in turn releases any OPC groups
' and items those groups may contain.
'
' This is a good catch all for your application but normally
' you should not rely on this call to cleanup your OPC connections.
' The best practice is to specifically release the OPC items from
' your groups then release the groups from your servers and finally
' disconnect the OPC servers from your application.  This will be
' demonstated in the various delete functions for each of the
' objects.
'
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    ' When exiting your application you should ensure that all
    ' groups and items are removed from the OPCServer objects
    ' then remove the OPCServer objects. The OPCServerClass
    ' contains a cleanup section in its terminate routine that
    ' will remove all groups and items from the class when the
    ' object is deleted. As we remove OPCServerClass object from
    ' the OPCServers collection this terminate function is called
    ' automatically removing We still need to remove the
    ' OPCServerGroups collection that contains the
    ' OPCGroupClasss objects ourselves.
    If OPCServers.Count <> 0 Then
        Dim a As Integer
        a = OPCServers.Count
        For i = 1 To OPCServers.Count
            With OPCServers
                ' Make sure we disconnect from server before we remove it
                Set Module1.SelectedOPCServer = .Item(a)
                Module1.SelectedOPCServer.DisconnectOPCServer
                .Remove (a)
                a = a - 1
            End With
        Next i
    End If
            
    ' Close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
    ' Save the current size of the application for the next load but
    ' only if the application is not in a minimized state.
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub


Private Sub mnuOption_Click()

End Sub

Private Sub tvTreeView_Collapse(ByVal Node As ComctlLib.Node)
    ' When a server is selected we set the Module1.SelectedOPCGroup
    ' to nothing and clear the lvListView control.
    Set Module1.SelectedOPCGroup = Nothing
    ' Clear the current view of tags if any.
    lvListView.ListItems.Clear
    
End Sub

' This sub is the primary focal point for user interaction with the
' OPC server and goup objects. Depending on where you click within the
' tree view you will be presented with differing context menus that will
' allow you to add new servers, add new groups, view the properties of
' these objects or delete them.  Clicking within this view also allows
' you to select different objects and change the data you see in the
' lvListView control.
'
Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' The SelectedNode var is used to hold the Node item returned from
    ' the HitTest
    Dim SelectedNode As Node
    ' When the selected node is group node this var will be used to
    ' hold the parent node of the selected group, ie. what server the group
    ' belongs to.
    Dim NodParent As Node
    ' Once a group and server are determined the NewGroupSelection is used
    ' to hold the selected group.  This is then testing against the
    ' Module1.SelectedOPCGroup variable.  If they differ then a new
    ' lvListView list will potentially be build by the GetNewItemList
    ' function.  This allows you to select different groups and see
    ' the items for these groups in the lvListView control.
    Dim NewGroupSelection As OPCGroupClass

    ' Used to temporary provide access to the group collection of
    ' the OPCServerClass object pointed to by Module1.SelectedOPCServer.
    Dim OPCServerGroupsCls As Collection



    ' Determine if the user has selected a node form the tree view.
    Set SelectedNode = tvTreeView.HitTest(X, Y)

    ' Right button actions select server and group object and
    ' display the various context menus.
    If (Button = vbRightButton) Then
        ' In all cases the "New Server Connection" menu selection
        ' is available.
        mnuTreeViewNewServer.Visible = True
        mnuTreeViewNewServer.Enabled = True
        
        ' If the user has selected a node as indicated by
        ' SelectedNode not being nothing then determine what type
        ' of node was selected.  This is where the Key name I mention
        ' in the start of this module really comes into play.  By
        ' knowing that all server nodes have the word "Server" as part
        ' of their Key and all group nodes have the word "Group" as part
        ' of their Key we can easily determine what type of node the user
        ' has selected.
        If Not SelectedNode Is Nothing Then
            ' Check for a server node
            If InStr(SelectedNode.Key, "Server") Then
                ' When we get to the AddSelectedOPCServer function
                ' you will see that I also use this same KEY within
                ' in the OPCServers collection to key the server.
                ' This make selection of the server from the collection
                ' as easy as what you see in the next three lines.
                ' In your application you can of course use any keying
                ' methodology you desire as long as their unique and
                ' you manage them.
                With OPCServers
                    Set Module1.SelectedOPCServer = .Item(SelectedNode.Key)
                End With
                
                ' When a server is selected we set the Module1.SelectedOPCGroup
                ' to nothing and clear the lvListView control.
                Set Module1.SelectedOPCGroup = Nothing
                lvListView.ListItems.Clear
                
                ' Now enable all of the options that the user can use when
                ' a server connection has been selected.
                mnuTreeViewNewGroup.Visible = True ' Add group allowed
                mnuTreeViewNewGroup.Enabled = True
                mnuTreeViewDelete.Visible = True    ' Deleted the Server
                mnuTreeViewDelete.Enabled = True
                mnuTreeViewProperties.Visible = True 'Display Server Properties
                mnuTreeViewProperties.Enabled = True
                mnuTreeViewNewItem.Visible = False   ' New Items can't be added at this time
                mnuTreeViewNewItem.Enabled = False
                
            ' If a server hasn't neen selected then it must be a group
            ' node.  We test the node key just to be sure.
            ElseIf InStr(SelectedNode.Key, "Group") Then
                ' Get a reference to the parent of group node.
                Set NodParent = SelectedNode.Parent
                ' Set the Selected Server to this parent easy stuff as
                ' long as the keys are managed and not lost.
                With OPCServers
                    Set Module1.SelectedOPCServer = .Item(NodParent.Key)
                End With
                
                ' Now that we know the server connection we need to access
                ' the OPCGroupClass collection contained in that
                ' OPCServerClass object.
                Set OPCServerGroupsCls = Module1.SelectedOPCServer.GetOPCServerGroupCollection
                ' The NewGroupSelection is simply used to test the newly
                ' selected group against a group that may already be selected.
                ' You'll notice right away that this use of the SelectedNode.Key
                ' is a little different from the simple use when selecting
                ' an OPC server.  You'll see why in the AddOPCGroupMain function
                ' below but here is the short answer.  When a node is added to the
                ' tree view control the Node.Key must be unique regardless
                ' of parent/child relationships or not.  As you'll see in the
                ' AddOPCGroupMain function the OPCServerClass.AddOPCGroup
                ' function returns a group key.  That key is specific
                ' to a single group and the server.
                Set NewGroupSelection = OPCServerGroupsCls.Item(SelectedNode.Key)
                
                ' Now that we have the OPCGroupClass object the user
                ' has selected we test to see if it is the currently
                ' selected group.  If its not then we need to update the
                ' lvListView with a new list of items if they exist.
                If (Not NewGroupSelection Is Module1.SelectedOPCGroup) Then
                    ' A new group has been selected so we need to dump the current
                    ' list view and load a new one.
                    Set Module1.SelectedOPCGroup = NewGroupSelection
                    ' Dump the current listview contents and try to load
                    ' the items from the newly selected group if any.
                    GetNewItemList
                End If
                          
                ' Enable/Disable those option for a group selection
                mnuTreeViewNewServer.Visible = False ' Disable New Server
                mnuTreeViewNewServer.Enabled = False
                mnuTreeViewNewGroup.Visible = False     'Disable New Group
                mnuTreeViewNewGroup.Enabled = False
                mnuTreeViewDelete.Visible = True        ' Enable delete of group
                mnuTreeViewDelete.Enabled = True
                mnuTreeViewProperties.Visible = True    ' Enable view of group properties
                mnuTreeViewProperties.Enabled = True
                mnuTreeViewNewItem.Visible = True       ' Enable add item.
                mnuTreeViewNewItem.Enabled = True
            End If
        Else ' If no node was selected but the user right clicked then
             ' display the popup with only the "New Server Connection" enabled
            mnuTreeViewNewGroup.Visible = False
            mnuTreeViewNewGroup.Enabled = False
            mnuTreeViewDelete.Visible = False
            mnuTreeViewDelete.Enabled = False
            mnuTreeViewProperties.Visible = False
            mnuTreeViewProperties.Enabled = False
            mnuTreeViewNewItem.Visible = False
            mnuTreeViewNewItem.Enabled = False
        End If
        
        ' Show the TreeView popup.
        frmMain.PopupMenu mnuTreeView
        
    ' The left button is used to select and OPC server connection
    ' or group only no popup is displayed.
    ElseIf (Button = vbLeftButton) Then
        If Not SelectedNode Is Nothing Then
            If InStr(SelectedNode.Key, "Server") Then
                ' Get the selected OPCServerClass object using the
                ' SelectedNode.Key which is the same as the server key
                ' in the OPCServers collection.
                With OPCServers
                    Set Module1.SelectedOPCServer = .Item(SelectedNode.Key)
                End With
                ' When a server is selected we set the Module1.SelectedOPCGroup
                ' to nothing and clear the lvListView control.
                Set Module1.SelectedOPCGroup = Nothing
                lvListView.ListItems.Clear
            ElseIf InStr(SelectedNode.Key, "Group") Then
                ' Get a reference to the parent of group node.
                Set NodParent = SelectedNode.Parent
                ' Set the Selected Server to this parent
                With OPCServers
                    Set Module1.SelectedOPCServer = .Item(NodParent.Key)
                End With
                
                ' See explanation in the right click portion above.
                Set OPCServerGroupsCls = Module1.SelectedOPCServer.GetOPCServerGroupCollection
                Set NewGroupSelection = OPCServerGroupsCls.Item(SelectedNode.Key)
                If (Not NewGroupSelection Is Module1.SelectedOPCGroup) Then
                    ' A new group has been selected so we need to dump the current
                    ' list view and load a new one.
                    Set Module1.SelectedOPCGroup = NewGroupSelection
                    GetNewItemList
                End If
            End If
        End If
    
    End If
End Sub


' This sub is invoked from the mnuTreeView popup context menu
' and simply displays the frmSelectOPCServer form.
'
Private Sub mnuTreeViewNewServer_Click()
    Load frmSelectOPCServer
    frmSelectOPCServer.Show
    ' Prevent the tree view from being selected while the server
    ' selection menu is displayed.
    '
    ' You'll see that I check this property in the lvListView_MouseUp
    ' as a simple way of preventing user input in the listview without
    ' using the .Enabled property of the ListView which cause the
    ' whole thing to grey out.
    tvTreeView.Enabled = False
End Sub


' This sub is invoked from the mnuTreeView popup context menu
' and simply displays the frmAddGroup form.
'
Private Sub mnuTreeViewNewGroup_Click()
    Load frmAddGroup
    frmAddGroup.Show
    ' Prevent the tree view from being selected while the server
    '  selection menu is displayed
    '
    ' You'll see that I check this property in the lvListView_MouseUp
    ' as a simple way of preventing user input in the listview without
    ' using the .Enabled property of the ListView which cause the
    ' whole thing to grey out.
    tvTreeView.Enabled = False
End Sub


' This sub is invoked from the mnuTreeView popup context menu
' and simply displays the frmAddItem form. This is only
' available when a group is selected from the tvTreeView control.
' Since this is the sane as the ListView action I use that sub.
'
Private Sub mnuTreeViewNewItem_Click()
    ' Simply call the list view handler here to add a new item to the
    ' seleted group.
    mnuListViewNewItem_Click
End Sub


' This sub is invoked from the mnuTreeView popup context menu
' and simply displays either the Server properties form or the
' Group properties form.
'
Private Sub mnuTreeViewProperties_Click()
    ' In this case we know that the only way this menu could
    ' have been invoked is when a treeview node has been selected
    ' so we can simply test the tvTreeView.SelectedItem.Key for
    ' either the "Server" portion or "Group" portion of the key.
    If InStr(tvTreeView.SelectedItem.Key, "Server") Then
        Load frmOPCServerProperties
        frmOPCServerProperties.Show
    ElseIf InStr(tvTreeView.SelectedItem.Key, "Group") Then
        Load frmOPCGroupProperties
        frmOPCGroupProperties.Show
    End If
    ' Prevent the tree view from being selected while the properties
    ' menu is displayed.
    '
    ' You'll see that I check this property in the lvListView_MouseUp
    ' as a simple way of preventing user input in the listview without
    ' using the .Enabled property of the ListView which cause the
    ' whole thing to grey out.
    tvTreeView.Enabled = False
End Sub


' Up to this point the subs called from the mnuTreeView popup menu have simply
' invoked other VB forms.  The delete menu selection has a direct effect
' upon the state of you OPC data.  The mnuTreeViewDelete_Click will either
' delete the selected OPCServerClass object from the OPCServers
' collection and treeview or delete the selected OPCGroupClass from the
' selected OPCServerClass object.  In both cases the TreeView and ListView
' will also be updated accordingly.  This must be done to keep the node keys
' in sync with the OPC server and group object keys.
'
Private Sub mnuTreeViewDelete_Click()
    On Error GoTo SkipDelete
    ' Is the selected node in the treeview a server object?
    If InStr(tvTreeView.SelectedItem.Key, "Server") Then
        ' When removing the entire server from the system
        ' we first need to remove the groups that are part
        ' of that server.  We can do this by using the
        ' OPCServerGroups collection that is part of the OPCServerClass object
        ' The Groups are removed in reverse order as the collection
        ' shrinks this is important since the size of the collection goes down
        ' and the numerical indexes are adjust as you delete objects from
        ' the collection.
        
        Dim OPCServerGroupCls As Collection
        Set OPCServerGroupCls = Module1.SelectedOPCServer.GetOPCServerGroupCollection
        
        ' If there are groups in this server remove them from the OPCServerClass
        ' object and from the treeview.
        If OPCServerGroupCls.Count <> 0 Then
            Dim i As Integer
            Dim a As Integer
            Dim GroupKeyToRemove As String
            a = OPCServerGroupCls.Count ' intialize the count used when removing the objects in reverse order.
            For i = 1 To OPCServerGroupCls.Count
                Set Module1.SelectedOPCGroup = OPCServerGroupCls.Item(a)
                ' Get the group object key as a string
                GroupKeyToRemove = Module1.SelectedOPCGroup.GetGroupKey
                ' Remove the object from the OPCServerClass object's
                ' collection of group objects.
                Module1.SelectedOPCServer.RemoveOPCGroup (GroupKeyToRemove)
                tvTreeView.Nodes.Remove (GroupKeyToRemove)
                Set Module1.SelectedOPCGroup = Nothing
                a = a - 1
            Next i
        End If
        
        ' Clear the list view for this group
        lvListView.ListItems.Clear
                
        ' Now that all the groups have been removed from the OPCServerClass
        ' object we can disconnect the server from our application.  This
        ' is done by calling the DisconnectOPCServer function of the
        ' OPCServerClass object.
        Module1.SelectedOPCServer.DisconnectOPCServer
        
        ' Now we remove the OPCServerClass object itself from the
        ' the OPCServers collection.
        With OPCServers
            .Remove (tvTreeView.SelectedItem.Key)
        End With
        ' Finally we remove the server connection from the treeview.
        tvTreeView.Nodes.Remove (tvTreeView.SelectedItem.Key)
        ' clear the SelelctedOPCServer pointer.
        Set Module1.SelectedOPCServer = Nothing
     
    ' Removing the group object from a server object is a little easier
    ' since we don't need to do any looips to remove a list of groups
    ' from the OPCServerClass object or treeview.
    
    ElseIf InStr(tvTreeView.SelectedItem.Key, "Group") Then
        Dim Result As Boolean
        ' Atempt to remove the selected OPCGroupClass object from the
        ' selected OPCServerClass object by calling the RemoveOPCGroup method
        ' of the OPCServerClass object.
        Result = Module1.SelectedOPCServer.RemoveOPCGroup(tvTreeView.SelectedItem.Key)
        
        ' If the remove was sucessful remove the group from the treeview
        ' and clear the listview control.
        If Result = True Then
            Set Module1.SelectedOPCGroup = Nothing
            tvTreeView.Nodes.Remove (tvTreeView.SelectedItem.Key)
            lvListView.ListItems.Clear
        End If
    End If
SkipDelete:
End Sub


' The GetNewItemList sub handles loading a new list view with OPC
' items from the selected OPCGroupClass object.  This occurs
' whenever the user selects a new group from the treeview.
'
Private Sub GetNewItemList()
    On Error Resume Next
    ' First clear the existing list view
    lvListView.ListItems.Clear
    LastTopItem = -1
    
    Dim OPCItemData As OPCItemClass
    Dim OPCGroupItemsCls As Collection
    ' Get pointer to the collection of items in the OPCGroupClass object.
    Set OPCGroupItemsCls = Module1.SelectedOPCGroup.GetOPCGroupItemsCollection
    
    ' If there are no items exit
    If OPCGroupItemsCls.Count = 0 Then
        GoTo SkipListUpdate
    End If
     
    Dim i As Integer
    
    For i = 1 To OPCGroupItemsCls.Count
        ' Normally we use the "Item Key" as string to specify the item we want
        ' from the collection but in this case we don't know the item keys
        ' explicitly sense some items may have been deleted.  To
        ' address this we can use the item index as a numeric index
        ' then pull the item key from the OPCItemData and use this as
        ' our list view key.  This demonstrates how both the string key
        '  and the numeric key of a collection can be used to address
        ' what otherwise would be a real key management issue.
        
        ' Get an OPCItemClass object from the items collection
        Set OPCItemData = OPCGroupItemsCls.Item(i)
        
        Dim ItemKey As String
        Dim itmX As ListItem
        Dim Quality As Long
        Dim ItemValue As Variant
        
        ' Build a unique strin key for the list view item using the actual
        ' item index number.
        ItemKey = "Item" + Str(OPCItemData.GetItemIndex)
        ' Add Node objects.
        Set itmX = lvListView.ListItems.Add(, ItemKey, OPCItemData.GetItemID)
            ItemValue = OPCItemData.GetItemValue(OPCItemDirect)
            ' When not processing an array value then simply display
            ' the value as it is in variant form.
            If Not IsArray(ItemValue) Then
                itmX.SubItems(1) = ItemValue
            Else
                ' If we have an array then we need to format the display
                ' by converting the value to a string with each value
                ' separated by commas.
                Dim b As Integer
                itmX.SubItems(1) = ""
                ' Regardless of the Option Base setting arrays returned
                ' from an OPC Server have a starting index of 0
                For b = 0 To UBound(ItemValue)
                    itmX.SubItems(1) = itmX.SubItems(1) + Str(ItemValue(b)) + ", "
                Next b
            End If
                
            ' Get the item quailty condition directly from the OPC server
            Quality = OPCItemData.GetItemQuality(OPCItemDirect)
            
            ' The "Data Access Automation Interface Standard" specification
            ' doesn't directly tell you the contents of the quailty field but
            ' in general if the field contains the value 192 of 0xC0 Hex the
            ' data can be considered OK.
            If Quality And &HC0 Then
               itmX.SubItems(2) = "Good" ' if quailty is 192 then OK
            Else
               itmX.SubItems(2) = "Bad " + Str(Quality) ' If the not 192 show Bad and value.
            End If
    Next i
              
SkipListUpdate:

End Sub


' Like the MouseUp handler for the TreeView, this sub handles all
' user interaction with the ListView control.  If the user right clicks
' in the lsit view they will be presented with a context popup. If the
' user has an item selected when the right button is pressed they will
' be able to add a new item, make the selected item active or inactive,
' write to the item, or delete it.  If the user right click without selecting
' an item they will be able to add a new item.
'
Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' As I mentioned above I use the enabled state of the tree view
    ' to determine if user input is allowed in the ListView.  I do
    ' this since at each point a new form will be displayed I disable
    ' the TreeView control.  This causes the treeview to become greyed out
    ' and prevents the user from invloking a popup from the tree view.
    ' If I uses the .enabled property of the ListView the whole view
    ' gets greyed out and looks pretty bad.  By monitoring the
    ' .enabled state of the treeview I can disble the listview input while
    ' keeping the display clear.
    If tvTreeView.Enabled = True Then
        ' First we need to make sure that a group is selected in the
        ' treeview control.
        If Not tvTreeView.SelectedItem Is Nothing Then
            If InStr(tvTreeView.SelectedItem.Key, "Group") Then
                ' A group is selected in the treeview if the right button
                ' is pressed show a popup.
                If (Button = vbRightButton) Then
                    ' The "New Item" selection is always allowed.
                    mnuListViewNewItem.Visible = True
                    mnuListViewNewItem.Enabled = True
                    
                    ' Is an item selected in the ListView
                    If Not lvListView.SelectedItem Is Nothing Then
                       Dim SelectedItem As ListItem
                       ' We use the hit test here to make sure the user
                       ' is directly on a specific item.
                       Set SelectedItem = lvListView.HitTest(X, Y)
                       If Not SelectedItem Is Nothing Then
                            Dim OPCGroupItemsCls As Collection
                            Set OPCGroupItemsCls = Module1.SelectedOPCGroup.GetOPCGroupItemsCollection
                            ' Set the SelectedOPCItem pointer
                            Set Module1.SelectedOPCItem = OPCGroupItemsCls.Item(Mid(SelectedItem.Key, InStr(SelectedItem.Key, " ")))
                            ' Display all Item options on the popup.
                            mnuListViewSetActive.Visible = True
                            mnuListViewSetActive.Enabled = True
                            mnuListViewSetInactive.Visible = True
                            mnuListViewSetInactive.Enabled = True
                            mnuListViewWriteValue.Visible = True
                            mnuListViewWriteValue.Enabled = True
                            mnuListViewDelete.Visible = True
                            mnuListViewDelete.Enabled = True
                        Else
                            ' The user is not directly on an item so show only
                            ' the "Add Item" selection.
                            mnuListViewSetActive.Visible = False
                            mnuListViewSetActive.Enabled = False
                            mnuListViewSetInactive.Visible = False
                            mnuListViewSetInactive.Enabled = False
                            mnuListViewWriteValue.Visible = False
                            mnuListViewWriteValue.Enabled = False
                            mnuListViewDelete.Visible = False
                            mnuListViewDelete.Enabled = False
                        End If
                    Else
                        ' The user is not directly on an item so show only
                        ' the "Add Item" selection.
                        mnuListViewSetActive.Visible = False
                        mnuListViewSetActive.Enabled = False
                        mnuListViewSetInactive.Visible = False
                        mnuListViewSetInactive.Enabled = False
                        mnuListViewWriteValue.Visible = False
                        mnuListViewWriteValue.Enabled = False
                        mnuListViewDelete.Visible = False
                        mnuListViewDelete.Enabled = False
                    End If
                
                ' Show the popup
                frmMain.PopupMenu mnuListView
                
                End If
            End If
        End If
    End If
End Sub


' The mnuListViewSetActive_Click directly changes the active state of
' the selected OPC item.  This is done by calling the .SetItemActiveState
' of the selected OPCItemClass object.
'
Private Sub mnuListViewSetActive_Click()
 
    ' Make sure an item is selected.
    If Not Module1.SelectedOPCItem Is Nothing Then
        ' Change its active state to True.  Normally you will want to call
        ' this function only once with out an associated call to make the item
        ' inactive.  This however is not a requirement and you can call this
        ' function on an item as many times as you like. The next call
        ' with the state set to False will still turn the item off.
        Module1.SelectedOPCItem.SetItemActiveState (True)
    End If
    
End Sub


' The mnuListViewSetInactive_Click directly changes the active state of
' the selected OPC item.  This is done by calling the .SetItemActiveState
' of the selected OPCItemClass object.
'
Private Sub mnuListViewSetInactive_Click()
    
    ' Make sure an item is selected.
    If Not Module1.SelectedOPCItem Is Nothing Then
        ' Change its active state to False.  Normally you will want to call
        ' this function only once with out an associated call to make the item
        ' active.  This however is not a requirement and you can call this
        ' function on an item as many times as you like. The next call
        ' with the state set to True will still turn the item on.
        Module1.SelectedOPCItem.SetItemActiveState (False)
    End If
    
End Sub


' This sub simply displays the frmWriteItem form.
'
Private Sub mnuListViewWriteValue_Click()
    'Module1.SelectedOPCItem will be set if this event occurs.
    Load frmWriteItem
    frmWriteItem.Show
    ' Prevent the tree view from being selected while the write item
    ' menu is displayed
    '
    ' You'll see that I check this property in the lvListView_MouseUp
    ' as a simple way of preventing user input in the listview without
    ' using the .Enabled property of the ListView which cause the
    ' whole thing to grey out.
    tvTreeView.Enabled = False
End Sub


' This sub directly removes an OPCIemClass object from the selected
' OPCGroupClass object.
'
Private Sub mnuListViewDelete_Click()
    On Error GoTo SkipItemDelete
      
    Dim Result As Boolean
    Dim Error As Long
    ' Attempt to remove the selected OPCItemClass object from the
    ' selected OPCGroupClass object by passing in the item key to the
    ' .RemoveOPCItem function of the class.
    Result = Module1.SelectedOPCGroup.RemoveOPCItem(lvListView.SelectedItem.Key, Error)
    ' If teh item was removed then remove it from the list view.
    If Result = True Then
        Set Module1.SelectedOPCItem = Nothing
        lvListView.ListItems.Remove (lvListView.SelectedItem.Key)
    End If
    
SkipItemDelete:
End Sub


' This sub is invoked from the mnuListView popup context menu
' and simply displays the frmAddItem form.
'
Private Sub mnuListViewNewItem_Click()
    Load frmAddItem
    frmAddItem.Show
    ' Prevent the tree view from being selected while the Add item
    ' menu is displayed
    '
    ' You'll see that I check this property in the lvListView_MouseUp
    ' as a simple way of preventing user input in the listview without
    ' using the .Enabled property of the ListView which cause the
    ' whole thing to grey out.
    tvTreeView.Enabled = False
End Sub


' The timer routine is used to update the contents of the list view display
' based on the contents data in the OPCItemClass objects attached to the
' selected OPCGroupClass object.  Normally we could allow the actual update
' of the contents of the OPCItemsClass object do the updating of the listview
' but the flashing would be really bad.  So I chose to use a timer tick
' to update the display. It still flashes but not a badly.  There are ways to
' prevent the flashing but they go beyond the scope of this example.
' The other little pain you'll notice is that you can't scroll the list
' unless you select an Item at the bottom of each page of the list.  I was
' not able to find a way to remove the selection from the item so the list
' could be scrolled easily.  When a listview item is selected the list view
' tries to get that value in the view.
'
Private Sub Timer1_Timer()
    On Error GoTo SkipDisplayUpdate
    
    Dim OPCGroupItemsCls As Collection
    
    ' Make sure a group is selected before any update occurs.
    If InStr(tvTreeView.SelectedItem.Key, "Group") Then
    
        Dim OPCGroupToUpdate As OPCGroupClass
        Dim OPCItemData As OPCItemClass
        Dim OPCServerGroupsCls As Collection
        
        ' Get the OPCItemClass object collection from the selected OPC group
        Set OPCServerGroupsCls = Module1.SelectedOPCServer.GetOPCServerGroupCollection
        Set OPCGroupToUpdate = OPCServerGroupsCls.Item(tvTreeView.SelectedItem.Key)
        Set OPCGroupItemsCls = OPCGroupToUpdate.GetOPCGroupItemsCollection
        
        ' If there aren't any items in the group skip the update
        If OPCGroupItemsCls.Count = 0 Then
            GoTo SkipDisplayUpdate
        End If
        
        ' only update those items that are within the display area of the
        ' list view.
        Dim itmX As ListItem
        ' Get the first Listitem in the ListView display
        Set itmX = lvListView.GetFirstVisible
         
        ' This should contain a valid ListItem but just in case testing
        ' it is always a good idea.
        If Not itmX Is Nothing Then
            Dim NumLinesDisplayed  As Integer
            ' This a hack, normally you would get the font metric for
            ' the system font, get the height and calculate the number
            ' of lines in the display area.  I didn't want to do this
            ' at the time so here's a rough estimate of the value.
            NumLinesDisplayed = (lvListView.Height / 214)
            
            Dim i As Integer
            Dim a As Integer
            Dim GroupItemIndex As Integer
            Dim OPCItemToUpdate As OPCItemClass
            Dim Quality As Long
            Dim ItemValue As Variant
            a = itmX.Index
            
            ' See if the user is scrolling up or down and change the
            ' selected ListView item to allow the view to move beyond the
            ' current items in the view.
            If LastTopItem <> a Then
                If LastTopItem <> -1 Then
                    Set lvListView.SelectedItem = itmX
                End If
                LastTopItem = a
            End If
                       
            For i = 1 To NumLinesDisplayed
                ' Get the ListItem key which is in the form "Item X"
                ' Since the OPCItemClass collection is key on a string
                ' value of just "X" we must remove the "Item" portion
                ' and conver thenumber back to a string.
                GroupItemIndex = Val(Mid(itmX.Key, InStr(itmX.Key, " ")))
                ' Now that we have the key as a string we can get the
                ' OPCItemClass object associated with this listview item.
                Set OPCItemToUpdate = OPCGroupItemsCls.Item(Str(GroupItemIndex))
                
                ' With the OPCItemClass object in hand we can update the list
                ' view item.  For both the ItemValue and the Quality value
                ' I use the Const OPCItemLocal which forces the OPCItemClass
                ' functions to return their local copies of the data value and
                ' quality.
                ItemValue = OPCItemToUpdate.GetItemValue(OPCItemLocal)
                
                ' If the ItemValue is not an array then simply use the Variant
                ' type to update the list box sub item 1 which the Value Heading.
                If Not IsArray(ItemValue) Then
                    itmX.SubItems(1) = ItemValue
                Else
                    ' If it is an array we need to turn the array into a string
                    ' that shows the values separated by commas.
                    Dim b As Integer
                    itmX.SubItems(1) = "" ' listview item string to start
                    ' Regardless of the Option Base setting arrays returned
                    ' from an OPC Server have a starting index of 0
                    For b = 0 To UBound(ItemValue)
                        ' Build the array display by looping through each value and appending it with a comma.
                        itmX.SubItems(1) = itmX.SubItems(1) + Str(ItemValue(b)) + ", "
                    Next b
                End If
                
                ' Get the item Quailty for this item.
                Quality = OPCItemToUpdate.GetItemQuality(OPCItemLocal)
                If Quality And &HC0 Then
                    itmX.SubItems(2) = "Good"
                Else
                    itmX.SubItems(2) = "Bad" + Str(Quality)
                End If
                            
                ' Get the next ListView item using its numeric index number.
                ' If we exceed the number of item available in the list
                ' display the error handler will automatically to us to the
                ' SkipDisplayUpdate label
                a = a + 1
                Set itmX = lvListView.ListItems.Item(a)
            Next i
            
      End If
        
    End If
SkipDisplayUpdate:
End Sub


' The AddSelectedOPCServerMain is called from the frmSelectOPCServer form.
' The function is housed here because of the dependecy it has on
' updating the TreeView control once a new server connection is added.
' This function also creates the actual working instance of the OPCServerClass
' object that will be used to contain your OPC groups and their associated
' items.  If the server connection is successfully made with the server
' the OPCServers collection will be updated to contain the new OPCServerClass
' object.  The TreeView will also be updated with this new server connection.
' Whether your OPC application has a user interface or is driver progrmatically,
' you will always start by making a connection to an OPC server first.
' Once the connection to the OPC server is made you can add groups and items
' to the connection and start putting OPC to use.  In this example this
' function handle both the class objects and the user interface objects.
' In your application you may not need the user interface code.
'
Sub AddSelectedOPCServerMain(OPCServerName As String)
    ' Create a OPCServerClass object that we will attempt to connect to the server
    ' if the connection is sucessful we will add this new OPCServerClass object to
    ' OPCServers collection.
    Dim OPCServer As OPCServerClass
    Set OPCServer = New OPCServerClass
    Dim Result As Boolean
    
    ' Create a unique key for this new server
    Dim SrvName As String
    SrvName = "Server" + Str(Module1.ServerIndex)
          
    ' Attempt a connection to the selected OPC Server
    Result = OPCServer.ConnectOPCServer(OPCServerName, SrvName, Module1.ServerIndex)
          
    ' Prepare for next server connection.
    Module1.ServerIndex = Module1.ServerIndex + 1
    
    ' If the connection was sucessful then add this new OPCServerClass
    ' object to the OPCServers collection and add this to the TreeView.
    If (Result = True) Then
           
        With OPCServers
           .Add OPCServer, SrvName
        End With
    
        ' Add Node objects.
        Dim nodX As Node    ' Declare Node variable.
        ' Add the new server as a root in the tree view
        Set nodX = fMainForm.tvTreeView.Nodes.Add(, , SrvName, OPCServerName)
        nodX.EnsureVisible
    
        ' Make this new server the selected server
        Set Module1.SelectedOPCServer = OPCServer
        ' Clear any selected group or item
        Set Module1.SelectedOPCGroup = Nothing
        Set Module1.SelectedOPCItem = Nothing
        ' Remove all items from the list view if there are any.
        lvListView.ListItems.Clear
    
    End If
    
    ' If the connect fails the DisplayOPC_COM_ErrorValue inside of the OPCServerClass
    ' will display the error.
    
    End Sub
    
    
' The AddOPCGroupMain function is called from the frmAddGroup form.
' The function is housed here because of the dependecy it has on
' updating the TreeView control once a new group object is added.
' Unlike the AddSelectedOPCServerMain function above this function does not
' actually create the OPCGroupClass object.  That is done in the
' OPCServerClass object you now have.  The OPCServerClass objects are handled
' directly by your application.  After that the OPCGroupClass object and
' OPCItemClass object are handled by the class modules.  You have access to the
' the collections within these object that house the groups or items but
' your application doesn't directly need to store these objects.
' You'll also notice that unlike the OPCServerClass object, we as the
' application don't determine the OPC group key instead the AddOPCGroup
' function of the OPCServerClass object returns a key to us.  This key
' can then be used in the list view as a key.  The key is unique for all
' server connection since it is built using the group index and the server
' index numbers.  Example group keys are "Group 1 1" for a group on
' OPCServerClass object with a server index of 1, and "Group 1 2" on an
' OPCServerClas object with a server index of 2.  In this example this
' function handle both the class objects and the user interface objects.
' In your application you may not need the user interface code.
'
Sub AddOPCGroupMain(ByVal GroupName As String, ByVal UpdateRate As Long, ByVal DeadBand As Single, ByVal ActiveState As Boolean)
    
    Dim GroupKey As String
               
    If Module1.SelectedOPCServer.AddOPCGroup(GroupName, UpdateRate, DeadBand, ActiveState, GroupKey) = True Then
        ' Add Node objects.
        Dim nodX As Node    ' Declare Node variable.
        ' Add the new server as a root in the tree view
        Set nodX = fMainForm.tvTreeView.Nodes.Add(Module1.SelectedOPCServer.GetOPCServerKey, tvwChild, GroupKey, GroupName) ' + Str(Module1.SelectedOPCServer.GetOPCServerIndex)
        nodX.EnsureVisible
    End If
    
    ' If the add group fails the DisplayOPC_COM_ErrorValue inside of the OPCServerClass
    ' will display the error.
    
End Sub


' The AddOPCItemMain function is called from the frmAddItem Form.
' The function is housed here because of the dependecy it has on
' updating the ListView control once a new group object is added.
' Unlike the AddSelectedOPCServerMain function above this function does not
' actually create the OPCItemClass object.  That is done in the
' OPCGroupClass object you now have.  Once you have a group added to your
' OPCServerClass object you can add items to it by calling the
' OPCGroupClass.AddOPCItem function.  The AddOPCItem function takes
' the ItemID, Datatype, and Active state as parameters. It also
' takes a ItemKey string to hold the returned ItemKey for use in the
' ListView control.  In this example this
' function handle both the class objects and the user interface objects.
' In your application you may not need the user interface code.
'
Function AddOPCItemMain(ByVal ItemID As String, ByVal DataTypeSelected As Integer, ByVal ActiveState As Integer)
    
    Dim ItemKey As String
    
    ' Make sure the SelectedOPCGroup pointer is valid
    If Not Module1.SelectedOPCGroup Is Nothing Then
        ' Attempt to add an OPC item to the selected group.
        If Module1.SelectedOPCGroup.AddOPCItem(ItemID, DataTypeSelected, ActiveState, ItemKey) = False Then
            AddOPCItemMain = False
            GoTo ErrorOnAdd
        End If
    End If
    
    Dim itmX As ListItem
    ' Add the new item to the ListView's ListItem objects.
    Set itmX = lvListView.ListItems.Add(, ItemKey, ItemID)
    itmX.SubItems(1) = "" 'Initialize to no value
    itmX.SubItems(2) = "Bad" ' Initialize to Bad quality
    
    ' Let the caller know the function worked.
    AddOPCItemMain = True
ErrorOnAdd:
End Function
    
    
    
    
    
'*****************************************************************
' The remainder of this module was code generated by the VB wizard
'*****************************************************************

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub


Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewUpdateDispaly_Click()
    Load frmItemUpdateInterval
    frmItemUpdateInterval.Show
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    'set the width
    If X < 3500 Then X = 2500
    If X > (Me.Width - 2500) Then X = Me.Width - 2500
    tvTreeView.Width = X
    imgSplitter.Left = X
    lvListView.Left = X + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40


        
    tvTreeView.Top = picTitles.Height
    lvListView.Top = tvTreeView.Top
    

    'set the height
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    

    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub

