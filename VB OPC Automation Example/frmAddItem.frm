VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAddItem 
   Caption         =   "Add Item"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Browsing 
      Caption         =   "Browsing"
      Height          =   2895
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   8415
      Begin VB.ComboBox AccessFilter 
         Height          =   315
         ItemData        =   "frmAddItem.frx":0000
         Left            =   6720
         List            =   "frmAddItem.frx":0010
         TabIndex        =   18
         Text            =   "Any"
         ToolTipText     =   "Select an access type to filter the list of items displayed"
         Top             =   360
         Width           =   1215
      End
      Begin ComctlLib.ListView lvItemView 
         Height          =   2055
         Left            =   3960
         TabIndex        =   14
         ToolTipText     =   "Select an item from this list by double clicking on the item"
         Top             =   720
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3625
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox DataTypeFilter 
         Height          =   315
         ItemData        =   "frmAddItem.frx":003C
         Left            =   5400
         List            =   "frmAddItem.frx":006F
         TabIndex        =   13
         Text            =   "Native"
         ToolTipText     =   "Select a datatype to filter the list of items displayed"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox ItemNameFilter 
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Text            =   "*"
         ToolTipText     =   "Enter a filter expression to reduce the number of item displayed Ex. Temp* , N* ,  *"
         Top             =   360
         Width           =   1335
      End
      Begin ComctlLib.TreeView tvBranchView 
         Height          =   2535
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Expand the branches of the browser tree to access individual item names"
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4471
         _Version        =   327682
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Access Filter"
         Height          =   255
         Left            =   6720
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Datatype Filter"
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Leaf Filter"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton Abort 
      Caption         =   "Done"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      ToolTipText     =   "Click here when you have finished adding tags"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton AddItem 
      Caption         =   "Add Item"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Click here to add a tag to the current group"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Validate 
      Caption         =   "Validate"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      ToolTipText     =   "Once you have entered an ItemID and Datatype click here to test the entry"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CheckBox ActiveState 
      Alignment       =   1  'Right Justify
      Caption         =   "Active State:"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Set the initial Active state of the item to be added"
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox DataType 
      Height          =   315
      ItemData        =   "frmAddItem.frx":00D4
      Left            =   960
      List            =   "frmAddItem.frx":0107
      TabIndex        =   1
      Text            =   "Native"
      ToolTipText     =   "Select the datatype that will be applied to this item"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox ItemID 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "Channel_1.Device_1.R0 (or your default ItemID string)"
      ToolTipText     =   "Enter a valid Item ID String for the server you are using, this default item is for KEPServerEX"
      Top             =   360
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Definition"
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   6855
      Begin VB.TextBox Report 
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "This box will report the results of the Validate Item button"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Item ID:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Data Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The Add OPC Item form allows the user to set the ItemID,
' DataType, and Active state for an OPC item that will added
' to the selected group. When this form is displayed the
' tvTreeView on the main form is disable to prevent the
' user from invoking an Add group or Add Server menu.
' This form will only be invoke while the user has an
' existing OPC server group selected.  This means that
' global variable SelectedOPCGroup and SelectedOPCServer
' will be set.
'
' The Add item form is preconfigured with a simple default
' ItemID string that will work in our KEPServerEX
' OPC Servers when running the Simdemo
' project in server.  If you are using a driver
' other than our Simulator or using a differnt OPC server
' altogether simply replace this default ItemID with a
' fully qualified item id for the driver or server being used.
' You can also browse the tag space of the server you are connected
' with to receive a list of item IDs.
'
' You'll also notice that I preselect a DataType of "Native" for
' the item.  Normally when you add an item to an OPC server
' group you will set the datatype to format that you desire for
' the item.  OPC and COM objects in general support many data
' type formats.  In some cases you may not know the format of the
' data for ItemID you are trying to register.  In these cases you can
' specify what is refered to as either Empty or Native.  When
' selecting a datatype of Native you are telling the OPC server
' to pick the default data type of the item for you.
'
' There is one possible problem with using the Native datatype
' and allowing the server to pick the data type.  Sometimes the
' default data type for an item is an Unsigned format.  At this
' time the Automation Wrapper can have problems with some
' Unsigned datatypes.  I have seen a problem with arrays of
' Unsigned values.  I have however successfully used Unsigned
' values of single items.  In general try to select a data type
' from those presented in the drop down Datatype control.
' For more reference on the datatype available for OPC items
' use the VB help system and search for "VarType Constants"
' You'll also notice that I have provided some array type
' selections. An array data type is built by adding the VB
' Const vbArray + (desired type), you'll see this in the
' DataType combobox's ItemData and on the VB help page.
'
' The Active state allows you to set whether or not the OPC
' item you intend to add will be actively acquiring data or
' idle.  The active state of an OPC item is an important tool
' in controlling how an OPC server is talking to you phsyical
' device.  The amount of data that an OPC server is gathering
' from a device depends on two factors, the number of actual
' OPC items added to the OPC server/groups and the number of
' those item that are marked as active.  This is an important
' consideration.  When designing your OPC client application
' you can control the operation of the OPC server by adding
' OPC items and deleting them as need.  This works but it
' forces the OPC server to do a great deal of memory allocation
' and releasing as well as potentially interupting the order
' of how data is being polled from the server.  A better way
' to control how much data is being acquired by the OPC server
' is to use the active state of the item.  By using the active
' state of the item you can essential turn an item off. When
' this occurs the OPC server will stop scanning that item from
' the physical device.  This of course will reduce the
' bandwidth requirements of the OPC server.  By using the
' active state of the item instead of adding and deleting the
' items you allow the OPC server to operate at peak performance.
'
' The AddItem form also demonstrates another important and
' useful component of the OPC inteface, that is the Validate
' method found on the OPCItems object of the Automation
' Wrapper.  The Validate method allows you to ask the OPC
' server if the ItemID, DataType, and Active State of the item
' you intend to add is valid.  Essentially the Validate
' method is just like calling the acutal AddItem mehtod but
' it doesn't actually add the item to the server.  The OPC
' server will return error status of the item.  This is a great
' way to let the users of your OPC client application known
' that the item they are entering is valid
'
' Just like the Add OPC Group form, the actual addition of the
' OPC item occurs over in frmMain since this is where the item
' will be added to the lvListView control.  The actual
' AddOPCItem function of the OPCGroupClass is shown in frmMain.
' Calling the ValidateOPCItem of the OPCGroupClass is shown
' here directly.
'
' This form also handles most of tag browsing code for this
' example application.  The browse interface is housed in the
' OPCBrowserClass.  Most of the actual work is done here since
' majority of the work is in managing the treeviiew and listview
' controls.
'
' I take a simplified approach to maintaining browser sync with the OPC
' server.  The browse interface provides a number of method that allow you to
' move up and down the browse tree of the OPC server.  You can use these
' methods to change what segment of the server's tag space you can access
' at a given level.  In this example I keep track of the various branches
' in the browse space and use this data to do a simple Move operation that
' always starts at the root of the server browse space.  In this way it
' would be pretty hard to get out of sync with the server.



Option Explicit
Option Base 1

' This object is returned from the OPCServerClass function
' GetServerBrowseObject. The OPCBrowserClass object is a relative thin
' wrapper for the Automation browse object.
'
Dim OPCBrowserClassObj As OPCBrowserClass

' The variable is used to generate a number of dummy treeview
' nodes that allow the tree to always have an expansion available.
' In this way we can get the data for each branch of the browse
' space when the user expands a branch. You will see the
' addition of the Dummy node each time a branch is collapsed.
' By the same token the dummy node is removed before the branch is
' completely expanded so the user never sees this phatom node.
'
Dim DummyNodeNum As Integer

' Simle place holder for the connected server key string since it is
' used in number of places in the browser code.
Dim ServerKey As String


' On form load we initialize the treeview control and listview controls
' to allow the user to expand the treeview by one level below the
' actual server name.
Private Sub Form_Load()

Dim ServerName As String

' Get the OPCBrowserClass object that will be used during this
' instance of the AddItem dialog.
'
Module1.SelectedOPCServer.GetServerBrowseObject OPCBrowserClassObj

' Get the Server Name that will act as the base item in the
' treeview control.  This first branch of the treeview is always displayed.
'
Module1.SelectedOPCServer.GetServerName ServerName

' Get the server key used to determine when the branch selected is the root
' ServerName branch and also to be used to build unique branch keys
' in the tree view as the browse space is expanded.
'
ServerKey = Module1.SelectedOPCServer.GetOPCServerKey

' Intiailize the var used to create the dummy nodes that allow the
' treeview to always have an expansion available for every branch.
'
DummyNodeNum = 1

' Set the list view column headers.
lvItemView.ColumnHeaders.Add , , "ItemID", lvItemView.Width
lvItemView.View = lvwReport ' Place list view text report mode
    
' If the server supports browsing initialize the
' treeview and listview controls.
If Not OPCBrowserClassObj Is Nothing Then
        ' Add Node objects.
        Dim nodX As Node    ' Declare Node variable.
        ' Add the new server as a root in the tree view
        Set nodX = tvBranchView.Nodes.Add(, , ServerKey, ServerName)
        nodX.EnsureVisible
        
        Dim Organization As Long
        OPCBrowserClassObj.GetBrowserOrganization Organization
        
        ' If the server is a flat browse space then populate the list view
        ' now since there won't be any other branches.
        If Organization = 2 Then
            OPCBrowserClassObj.ShowLeafs False
            Dim i As Integer
            For i = 1 To OPCBrowserClassObj.GetItemCount
                Dim ItemName As String
                OPCBrowserClassObj.GetItemName ItemName, i
                lvItemView.ListItems.Add i, i, ItemName
            Next i
        Else
            ' Add a dummy node that will allow the first branch to
            ' be expanded. We use the phrase "KEPwareDummyNode" plus the DummyNodeNum
            ' to generate a dummy node that will always be unique.  The
            ' use the phrase "KEPwareDummyNode" is crucial since this string is
            ' used to identify the dummy node so it can be removed as the
            ' branch is expanded.  When you look at the treeview expand method
            ' you will see the dummy node is removed since it is known to be
            ' the first and only child when a branch is collapsed.  The
            ' dummy node is placed back into the treeview at the specific
            ' level when the branch is collapse after all other children
            ' for that brach have been removed.  In this way the treeview
            ' is always shown with little expansion + signs.
            '
            tvBranchView.Nodes.Add ServerKey, tvwChild, "KEPwareDummyNode" + Str(DummyNodeNum), ""
        End If
End If

DummyNodeNum = DummyNodeNum + 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Free the memory of the OPCBrowserClass object.
    Set OPCBrowserClassObj = Nothing
End Sub

' This first section deals with the Item Definition portion of the
' dialog.


' When the user clicks the Add Item button the ItemID,
' DataType, and Active state of teh item will be passed to
' the AddOPCItem function in the frmMain form.  With the browser
' you can add multiple tags to the group by duble clicking on the
' desired tag in the browsers list view and then clicking the Add
' button.
'
Private Sub AddItem_Click()
    
    Dim DataTypeSelected As Integer
    
    ' Make sure the user makes a specific selection for the
    ' data type
    If DataType.ListIndex = -1 Then
        DataTypeSelected = 0
    Else
        DataTypeSelected = DataType.ItemData(DataType.ListIndex)
    End If
    
    ' Call the AddOPCItem function of the frmMain and see if
    ' this new item can be added to the OPC group.  The actual
    ' call to the AddOPCItem of the OPCGroupClass can be seen in
    ' this frmMain form function.  Normally you would be calling
    ' the AddOPCItem function of the OPCGroupClass directly.
    
    If fMainForm.AddOPCItemMain(ItemID.Text, DataTypeSelected, ActiveState.Value) = False Then
        MsgBox "The OPC Server has rejected the item please check the Item ID and data type selections"
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
    frmAddItem.Show
End Sub


' This simply clears the Report.Text string when a
' change to this control occurs.  This is done to allow
' the validate to show a new report string.
'
Private Sub ItemID_Change()
    Report.Text = ""
End Sub

' This is just a cute way of clearing the contents
' the Report.Text control so we can provide a new
' validate response.
'
Private Sub Validate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
    If Button = vbLeftButton Then
        ' Let the user know we are about to test the item.
        Report.Text = "Testing Item"
    End If
End Sub

' This event handler will call the actual ValidateOPCItem
' function found in the OPCGroupClass object.  The
' instance of the OPCGroupClass used is contained in the
' variable SelectedOPCGroup of Module1.  In this example
' you can count on that variable being set at this point.
' In general it's very important to make sure that you
' use the appropriate OPCGroupClass object when attempting
' to either Validate or Add an OPC item to the server.  If you
' use the wrong instance of the OPCGroupClass object the
' Validate or Add functions may fail or worse you'll add an
' item to the wrong group.
'
Private Sub Validate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim DataTypeSelected As Integer
    Dim Error As Long
    
    If Button = vbLeftButton Then
        ' Make sure the user makes a specific selection for the
        ' data type
        If DataType.ListIndex = -1 Then
            DataTypeSelected = 0
        Else
            DataTypeSelected = DataType.ItemData(DataType.ListIndex)
        End If
        
        ' Although the Module1.SelectedOPCGroup is set at this
        ' point it's always good practice test it before we use it.
        If Not Module1.SelectedOPCGroup Is Nothing Then
            ' Call the ValidateOPCItem function of the OPCGroupClass
            If Module1.SelectedOPCGroup.ValidateOPCItem(ItemID.Text, DataTypeSelected, ActiveState.Value, Error) = False Then
                ' The OPC Server didn't like something about the
                ' ItemID, DataType or Active state.
                Report.Text = "Error - Check ItemID or Datatype"
            Else
                ' Item has been accepted as given.
                Report.Text = "ItemID and Datatype Accepted"
            End If
        End If
    End If
End Sub


' Simple cancel of the AddItem form
'
Private Sub Abort_Click()
    ' Reenable the tree view
    fMainForm.tvTreeView.Enabled = True
    Unload Me
End Sub


' This simply clears the Report.Text string when a
' change to this control occurs.  This is done to allow
' the validate to show a new report string.
'
Private Sub ActiveState_Click()
    Report.Text = ""
End Sub

' The Expand operation will attempt to display any branches that
' the server may contain at the current position of the expansion.
' The current position is established each time a branch of the
' treeview is expanded.  The treeview can always be expand due to
' the dummy node that is added to the control at each new branch.
' This cna be seen in the Collapse event.
'
Private Sub tvBranchView_Expand(ByVal Node As Node)

    ' Make sure the server support browsing.
    If Not OPCBrowserClassObj Is Nothing Then
        Dim i As Integer
        Dim ItemName As String
        Dim DataTypeSelected As Integer
        Dim AccessSelected As Long

    
        ' Remove the dummy node that we know will always exist
        ' when a branch is expanded.
        If Node.Children Then
            ' Make sure that the child node is the dummy node and
            ' remove it.
            If InStr(Node.Child.Key, "KEPwareDummyNode") Then
                tvBranchView.Nodes.Remove (Node.Child.Key)
            End If
        End If
        
        ' The node key is built using the name of each branch
        ' of the actual server browse space.  This makes moving to
        ' a particular branch easy.  The node key is built below in
        ' this function.
        '
        SetBrowsePosition Node.Key
       
        ' Make sure that the Name filter is set to empty
        ' before we call to get any branches.
        OPCBrowserClassObj.SetFilter ""
        
        ' Tell the OPC Server to load it's internal item collection
        ' with Branch items.
        OPCBrowserClassObj.ShowBranches
        
        ' Now add each new branch if any under the current branch.
        ' Also add a dummy node to these new branches so they have
        ' an expansion available as well.
        For i = 1 To OPCBrowserClassObj.GetItemCount
            DummyNodeNum = DummyNodeNum + 1
            ' Get a branch name that exist under the current position of the
            ' server.
            OPCBrowserClassObj.GetItemName ItemName, i
            ' Load this name as a child to the current node of our
            ' treeview control.  As you can see the node key of the new child
            ' is built using the current node key of the parent + the
            ' ASCII character CHR(2) + the new itme name.  In this way a
            ' nodes key will contain the full path to an item as
            ' you browse down the tree.
            ' A node key might appear as
            ' "Server 0\02Channel_1\02Device_1" In this case the \02 signifies the
            ' CHR(2) character.  I use the CHR(2) character as a delimiter
            ' between each branch name.  This in turn is used in the SetBrowsePosition
            ' function to find each of these branch names and move down the browse tree.
            ' I use the CHR(2) simply because it is a non printable character
            ' that in all likely hood will not appear in the name of a branch
            ' in a server.  The Server Key is applied to the first branch
            ' in the treeview and is done in the Form Load.
            '
            tvBranchView.Nodes.Add Node.Key, tvwChild, Node.Key + Chr(2) + ItemName, ItemName
            
            ' Add a dummy node to these new child branches that will allow
            ' them to be expanded when the user clicks on them.  In this case
            ' the relative parameter is the node key of the child we jsut added
            ' to the parent branch.
            tvBranchView.Nodes.Add Node.Key + Chr(2) + ItemName, tvwChild, "KEPwareDummyNode" + Str(DummyNodeNum), ""
        Next i
    End If

End Sub

' The SetBrowsePosition is my little short cut function that
' takes a node.key string and parses out the various branches of the node
' and then performs a series of MoveDown operations to get to the
' desired branch level.  The function always starts at the root of
' the server before it moves down so getting out of sync with the server
' should be hard to do in practice.
'
Private Sub SetBrowsePosition(ByVal Position As String)
    Dim NormalPosition As String
    Dim NextBranch As Integer
    Dim LastBranch As Integer
    Dim FindBranch As String
    
    ' Init the indexes used to parse the Node.Key string.
    NextBranch = 1
    LastBranch = 1
    ' Establish the character that will be used to delimit each
    ' branch string within the Node.Key.
    FindBranch = Chr(2)
    
    ' Start at the roo of the server's browse space.  This helps to keep
    ' this application in sync with the server's browse tree and position.
    OPCBrowserClassObj.MoveToRoot
    
    ' See is the key is just the server root name in which case we do nothing
    ' since we are already at the root.
    If Len(Position) > Len(ServerKey) Then
        ' First step remove the Server Key from the first part of the Node.Key
        NormalPosition = Mid(Position, (Len(ServerKey) + 2))
        
        ' Now loop through each branch name in the NormalPositon string
        ' and MoveDown to that new branch as we go.
        Do While (NextBranch < Len(NormalPosition) And NextBranch <> 0)
            NextBranch = InStr((NextBranch + 1), NormalPosition, FindBranch)
            If NextBranch <> 0 Then
                OPCBrowserClassObj.MoveDown (Mid(NormalPosition, LastBranch, NextBranch - LastBranch))
                LastBranch = NextBranch + 1
            Else
                If LastBranch = 1 Then
                    OPCBrowserClassObj.MoveDown (NormalPosition)
                Else
                    OPCBrowserClassObj.MoveDown (Mid(NormalPosition, LastBranch))
                End If
            End If
        Loop
    End If

End Sub

' This function will go an get the current tags available at this
' branch of the server's browse tree. This is done using the
' ShowLeafs function of the browser interface.  Unlike the
' browsing for branches using the ShowBranches function, we now
' want to use the three methods available to filter tags(leafs) from
' the lsit returned  by the server.  By default these filters are
' all disabled or in a show everything state.
'
Private Sub tvBranchView_Click()

    Dim SelectedNode As Node
    
    Set SelectedNode = tvBranchView.SelectedItem
    
    If Not SelectedNode Is Nothing Then
         Dim i As Integer
         Dim ItemName As String
         Dim DataTypeSelected As Integer
         Dim AccessSelected As Long
         
         ' Move to the current position of the treeview.  We need to do
         ' this each time since the use could move from one branch in
         ' the tree to a completely different branch not simple down one
         ' or up one.
         SetBrowsePosition SelectedNode.Key
        
         ' Clear any items that may currently be in the list view control.
         lvItemView.ListItems.Clear
         
         ' Set the Item name filter using the string found in the
         ' ItemNameFilter cotnrol.  By default this is "*" which
         ' most server take to mean everything.  You can use
         ' other wildcards such as ? in addition to *.  Using these
         ' Wildcard you can do some fairly fancy filtering. An example
         ' might be "Pressure_???_T*"  In this case and items with the
         ' the name "Pressure_" then three fields of anything then
         ' "_T" anything would be return when we call ShowLeafs.
         '
         OPCBrowserClassObj.SetFilter ItemNameFilter.Text
         
         ' Set the datatype filter.  Like the item name filter this
         ' filter allows you to reduce the number of items returned
         ' in the ShowLeafs function call based on the data type of the
         ' item.  In this case we need to test the index of the DataTypeFilter
         ' combo box control since the user may not have selected anything
         ' yet in which case we will use the "Native" data type which tells
         ' the server to return all datatypes.
         If DataTypeFilter.ListIndex = -1 Then
             DataTypeSelected = 0
         Else
             DataTypeSelected = DataTypeFilter.ItemData(DataTypeFilter.ListIndex)
         End If
         OPCBrowserClassObj.SetDataTypeFilter DataTypeSelected
         
         ' Set the access type filter.  Like the item name filter this
         ' filter allows you to reduce the number of items returned
         ' in the ShowLeafs function call based on the read write access
         ' of the item.  In this case we need to test the index of the
         ' AccessFilter combo box control since the user may not have
         ' selected anything yet in which case we will use the
         ' access type of "0" which tells the server to return all
         ' items regardless of access type.
         If AccessFilter.ListIndex = -1 Then
             AccessSelected = 0
         Else
             AccessSelected = AccessFilter.ItemData(AccessFilter.ListIndex)
         End If
         OPCBrowserClassObj.SetAccessFilter AccessSelected
         
         ' Finally show the new items based on the filters set above.
         OPCBrowserClassObj.ShowLeafs
         ' Add the new item if anny to the listview control.
         For i = 1 To OPCBrowserClassObj.GetItemCount
             OPCBrowserClassObj.GetItemName ItemName, i
             lvItemView.ListItems.Add , "Item" + Str(i), ItemName
         Next i
    End If
        
End Sub

' When the branch of the treeview control is collapsed we must remove
' all the childreen under that branch and replace them with a single
' dummy node to allow the branch to be reexpanded if the user desires.
Private Sub tvBranchView_Collapse(ByVal Node As Node)
    If Node.Children > 0 Then
        RemoveChildren Node, True
    End If
    DummyNodeNum = DummyNodeNum + 1
    ' Add the dummy node back into the treeview control under the parent
    ' node that has just been collapsed.
    tvBranchView.Nodes.Add Node.Key, tvwChild, "KEPwareDummyNode" + Str(DummyNodeNum), ""
End Sub

' The RemoveChildren function removes all of the nodes/branches under a
' supplied parent node.  This function is used recusively to remove
' other node below each branch.  The Root paramter is set to true
' by the main caller to signify that the top most node is not to be
' deleted.  Once this function is called recursively the Root flag is
' set false allowing the parent of each subsequent level to be deleted
' as the children are removed.
Function RemoveChildren(ByVal Node As Node, ByVal Root As Boolean)
    Dim i As Integer
    Dim nodecnt As Integer
    Dim test As String
    
    test = Node.Key
    nodecnt = Node.Children
    If nodecnt > 0 Then
        For i = 1 To nodecnt
            test = Node.Child.Key
            RemoveChildren Node.Child, False
        Next i
        
        If Not Root Then
            tvBranchView.Nodes.Remove Node.Key
        End If
    Else
        tvBranchView.Nodes.Remove Node.Key
    End If
End Function

' Once a list of tags is dislpayed in the ListView control they can be loaded
' into the ItemID text control by double clicking on the tag name.
' Once in the ItemID control you can add to the tag to the current group
' by clicking the Add button on the form.  This does not close
' the dialog however allowing you to continue to add more tags as needed.
'
Private Sub lvItemView_DblClick()
    ItemID.Text = OPCBrowserClassObj.GetItemID(lvItemView.SelectedItem.Text)
End Sub

' The follwing function deal with the manipulation of the three
' methods of filter the the items names(tags) available from the server
' for a given branch.

' The DataTypeFilter allows you to reduce the number of tags displayed
' based on the data type.  The other two fitler mehtods must also be applied
' here before the ShowLeafs function is called to reload the tags
' in the lsitview control.  This operation occurs everytime you change the
' the data type filter on the dialog.
'
Private Sub DataTypeFilter_Click()
    Dim i As Integer
    Dim ItemName As String
    Dim DataTypeSelected As Integer
    Dim AccessSelected As Long
    
    ' Remove all of the current items from the lsit view.
    lvItemView.ListItems.Clear
    
    ' Set the Item name filter using the string found in the
    ' ItemNameFilter cotnrol.  By default this is "*" which
    ' most server take to mean everything.  You can use
    ' other wildcards such as ? in addition to *.  Using these
    ' Wildcard you can do some fairly fancy filtering. An example
    ' might be "Pressure_???_T*"  In this case and items with the
    ' the name "Pressure_" then three fields of anything then
    ' "_T" anything would be return when we call ShowLeafs.
    '
    OPCBrowserClassObj.SetFilter ItemNameFilter.Text
    
    ' Set the datatype filter.  Like the item name filter this
    ' filter allows you to reduce the number of items returned
    ' in the ShowLeafs function call based on the data type of the
    ' item.  In this case we need to test the index of the DataTypeFilter
    ' combo box control since the user may not have selected anything
    ' yet in which case we will use the "Native" data type which tells
    ' the server to return all datatypes.
    If DataTypeFilter.ListIndex = -1 Then
        DataTypeSelected = 0
    Else
        DataTypeSelected = DataTypeFilter.ItemData(DataTypeFilter.ListIndex)
    End If
    OPCBrowserClassObj.SetDataTypeFilter DataTypeSelected
    
    ' Set the access type filter.  Like the item name filter this
    ' filter allows you to reduce the number of items returned
    ' in the ShowLeafs function call based on the read write access
    ' of the item.  In this case we need to test the index of the
    ' AccessFilter combo box control since the user may not have
    ' selected anything yet in which case we will use the
    ' access type of "0" which tells the server to return all
    ' items regardless of access type.
    If AccessFilter.ListIndex = -1 Then
        AccessSelected = 0
    Else
        AccessSelected = AccessFilter.ItemData(AccessFilter.ListIndex)
    End If
    OPCBrowserClassObj.SetAccessFilter AccessSelected
    
    ' Finally show the new items based on the filters set above.
    OPCBrowserClassObj.ShowLeafs
    ' Add the new item if anny to the listview control.
    For i = 1 To OPCBrowserClassObj.GetItemCount
        OPCBrowserClassObj.GetItemName ItemName, i
        lvItemView.ListItems.Add , "Item" + Str(i), ItemName
    Next i
End Sub

' The ItemNameFilter_Change allows you to reduce the number of tags displayed
' based on the name of the item.  The other two fitler mehtods must also be applied
' here before the ShowLeafs function is called to reload the tags
' in the lsitview control.  This operation occurs everytime you change the
' the item name filter on the dialog.
'
Private Sub ItemNameFilter_Change()
    Dim i As Integer
    Dim ItemName As String
    Dim DataTypeSelected As Integer
    Dim AccessSelected As Long
    
    ' Remove all of the current items from the lsit view.
    lvItemView.ListItems.Clear
    
    ' Set the Item name filter using the string found in the
    ' ItemNameFilter cotnrol.  By default this is "*" which
    ' most server take to mean everything.  You can use
    ' other wildcards such as ? in addition to *.  Using these
    ' Wildcard you can do some fairly fancy filtering. An example
    ' might be "Pressure_???_T*"  In this case and items with the
    ' the name "Pressure_" then three fields of anything then
    ' "_T" anything would be return when we call ShowLeafs.
    '
    OPCBrowserClassObj.SetFilter ItemNameFilter.Text
    
    ' Set the datatype filter.  Like the item name filter this
    ' filter allows you to reduce the number of items returned
    ' in the ShowLeafs function call based on the data type of the
    ' item.  In this case we need to test the index of the DataTypeFilter
    ' combo box control since the user may not have selected anything
    ' yet in which case we will use the "Native" data type which tells
    ' the server to return all datatypes.
    If DataTypeFilter.ListIndex = -1 Then
        DataTypeSelected = 0
    Else
        DataTypeSelected = DataTypeFilter.ItemData(DataTypeFilter.ListIndex)
    End If
    OPCBrowserClassObj.SetDataTypeFilter DataTypeSelected
    
    ' Set the access type filter.  Like the item name filter this
    ' filter allows you to reduce the number of items returned
    ' in the ShowLeafs function call based on the read write access
    ' of the item.  In this case we need to test the index of the
    ' AccessFilter combo box control since the user may not have
    ' selected anything yet in which case we will use the
    ' access type of "0" which tells the server to return all
    ' items regardless of access type.
    If AccessFilter.ListIndex = -1 Then
        AccessSelected = 0
    Else
        AccessSelected = AccessFilter.ItemData(AccessFilter.ListIndex)
    End If
    OPCBrowserClassObj.SetAccessFilter AccessSelected
    
    ' Finally show the new items based on the filters set above.
    OPCBrowserClassObj.ShowLeafs
    ' Add the new items if any to the listview control.
    For i = 1 To OPCBrowserClassObj.GetItemCount
        OPCBrowserClassObj.GetItemName ItemName, i
        lvItemView.ListItems.Add , "Item" + Str(i), ItemName
    Next i
End Sub


Private Sub AccessFilter_Click()
    Dim i As Integer
    Dim ItemName As String
    Dim DataTypeSelected As Integer
    Dim AccessSelected As Long
    
    lvItemView.ListItems.Clear
    
    ' Set the Item name filter using the string found in the
    ' ItemNameFilter cotnrol.  By default this is "*" which
    ' most server take to mean everything.  You can use
    ' other wildcards such as ? in addition to *.  Using these
    ' Wildcard you can do some fairly fancy filtering. An example
    ' might be "Pressure_???_T*"  In this case and items with the
    ' the name "Pressure_" then three fields of anything then
    ' "_T" anything would be return when we call ShowLeafs.
    '
    OPCBrowserClassObj.SetFilter ItemNameFilter.Text
    
    ' Set the datatype filter.  Like the item name filter this
    ' filter allows you to reduce the number of items returned
    ' in the ShowLeafs function call based on the data type of the
    ' item.  In this case we need to test the index of the DataTypeFilter
    ' combo box control since the user may not have selected anything
    ' yet in which case we will use the "Native" data type which tells
    ' the server to return all datatypes.
    If DataTypeFilter.ListIndex = -1 Then
        DataTypeSelected = 0
    Else
        DataTypeSelected = DataTypeFilter.ItemData(DataTypeFilter.ListIndex)
    End If
    OPCBrowserClassObj.SetDataTypeFilter DataTypeSelected
    
    ' Set the access type filter.  Like the item name filter this
    ' filter allows you to reduce the number of items returned
    ' in the ShowLeafs function call based on the read write access
    ' of the item.  In this case we need to test the index of the
    ' AccessFilter combo box control since the user may not have
    ' selected anything yet in which case we will use the
    ' access type of "0" which tells the server to return all
    ' items regardless of access type.
    If AccessFilter.ListIndex = -1 Then
        AccessSelected = 0
    Else
        AccessSelected = AccessFilter.ItemData(AccessFilter.ListIndex)
    End If
    OPCBrowserClassObj.SetAccessFilter AccessSelected
    
    ' Finally show the new items based on the filters set above.
    OPCBrowserClassObj.ShowLeafs
    ' Add the new items if any to the listview control.
    For i = 1 To OPCBrowserClassObj.GetItemCount
        OPCBrowserClassObj.GetItemName ItemName, i
        lvItemView.ListItems.Add , "Item" + Str(i), ItemName
    Next i

End Sub





