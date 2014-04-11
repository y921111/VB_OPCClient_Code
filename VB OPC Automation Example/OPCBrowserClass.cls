VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OPCBrowserClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The OPCBrowserClass object houses all of the functionallity used
' to interact with an OPC server's tag browser.  This OPCBrowserClass
' is used here largely to wrap the functions supplied by the Automation
' interface with error handlers.  Unlike the OPCServerClass,
' OPCGroupClass, and OPCItemClass, the OPCBrowserClass simply provides
' consistancy in the of the error handling and rounds out the classes
' supplied by this example.  As is the case with any object based
' design using this class module will allow you to modify the browser
' interface if needed.  I have not wrapped all of the function supplied
' by the OPCBrowser object but adding others if needed is easy.
'
' In this example the frmAddItem form contains all of the real meat
' for actually browsing the tag space of an OPC Server.  The main code
' is housed there due largely to the fact that most of what needs to be
' be done when browsing is the management of the Treeview and Listview
' controls used to display the tags.

Option Explicit
Option Base 1

' This is the actuial Automation Interface browser object this is
' wrapped by this class.
'
Dim OPCBrowserObject As OPCBrowser


' This function simply sets the OPCBrowse object that will be used by
' this instance of the OPCBrowserClass object.  This is normally done
' by using the OPCServerClass function GetServerBrowseObject that
' should be used to nromally get this object class.  You can however
' create an OPCBrowserClass object yourself if you have an OPCServer
' object upon which you can use the .CreateBrowser method and pass the.
' resulting OPCBrower object to this function.
'
Function SetBrowserObject(ByVal OPCBrowseObj As OPCBrowser)
    Set OPCBrowserObject = OPCBrowseObj
End Function

' This function allows you to determine the format of the browse space
' available from the connected OPC Server.  There are two possible
' types of browse space your application may encounter.  There
' is Flat(OPCFlat = 2) and Hierarchial(OPCHierarchical = 1}. Most
' OPC Servers you will emcounter by default support a Hierarchial
' browse space.  Hierarchial means that the tags or items available
' in the server will be organized into a tree like structure with
' branches and leafs.  This organiztion is the same as what you
' see in your Window's Explorer where you have directories and files.
' In Hierarchial mode the your code must be able to move up and down
' the branches of the tree inorder to access the various tags found at
' each level in the tree.
'
' In a Flat browse space the server returns all of the items/tags
' at a single level.  Just because the server returns all of the items
' at a single level doesn't mean the the item IDs for all of item will
' be simple references.  The item IDs may still contain complex grouping
' information that your application may find useful.
'
' In this example the frmAddItem form contains all of the actual
' code that does the real work of moving around in the browse space
' of of the OPC Server.
'
Function GetBrowserOrganization(ByRef Organization As Long)
'Set error handling for OPC Function
    On Error GoTo ShowOPCGetBrowserOrganizationError
    
    Organization = OPCBrowserObject.Organization
    GetBrowserOrganization = True
    GoTo SkipOPCGetBrowserOrganizationError
    
ShowOPCGetBrowserOrganizationError:
    Call DisplayOPC_COM_ErrorValue("GetBrowserOrganization", Err.Number)
    GetBrowserOrganization = False
SkipOPCGetBrowserOrganizationError:
    
End Function

' The next two functions are used to move around in the browser space
' of the OPC Server.  When browsing an OPC server you must tell the
' server where to go on it's browse tree.  This is down using the
' following functions.  By default you start at the root of the
' OPC Server's browse space.  You then use the function ShowBranches
' and GetItemCount to determine if there are any other branches
' available from this point in the tree.  If there are more branches
' available you can use the MoveDown function and specify a branch
' to move down to.  Once you move down a level you will need to call the
' ShowBranches function again to see if there are branches at this new
' level of the OPC server browse space.  These steps are repeated until
' you reach the end of a paticular brance.  Additionally as you move to
' each branch including the root you will also be calling ShowLeafs
' to determine of there are any tags available at a given branch.
'
' In this exmaple I use only the MoveToRoot and MoveDown functions
' to browse the server.  In the frmAddItem form you will see that
' I keep track of the branch structure as you navigate the browse
' space.  This allows me to simply move to root and then do a series
' of MoveDown operations to get to the desired branch.  The OPCBrowser
' object also supports a MoveUp and MoveTo operation.  I don't use
' either in this example.

' The MoveToRoot function will return the OPC Server's browse position
' to the root of the browse space.  This makes staying in sync with the
' the OPC Server's position in the browse space easy. See the
' frmAddItem form for more on the use of this function.
'
Function MoveToRoot()
'Set error handling for OPC Function
    On Error GoTo ShowOPCMoveToRootError
    
    OPCBrowserObject.MoveToRoot
    MoveToRoot = True
    GoTo SkipOPCMoveToRootError
    
ShowOPCMoveToRootError:
    Call DisplayOPC_COM_ErrorValue("MoveToRoot", Err.Number)
    MoveToRoot = False
SkipOPCMoveToRootError:

End Function

' The MoveDown function moves the OPC Server's browse position down
' to the branch specified by the string "Branch".  When passing a
' branch name to this function you should try to use the branch name
' exactly as it was given to you by the ShwoBranches function.  This
' will insure that you don't give the server an improper name and get
' the OPC Server's browse space out of sync with you own.  See the
' frmAddItem form for more on the use of this function.
'
Function MoveDown(ByVal Branch As String)
'Set error handling for OPC Function
    On Error GoTo ShowOPCMoveDownError
    
    OPCBrowserObject.MoveDown (Branch)
    MoveDown = True
    GoTo SkipOPCMoveDownError
    
ShowOPCMoveDownError:
    Call DisplayOPC_COM_ErrorValue("MoveDown", Err.Number)
    MoveDown = False
SkipOPCMoveDownError:

End Function

' The MoveUp function move the OPC Server's browse position up
' one level from the current position.  While you can use this
' function in conjunction with the MoveDown to navigate the
' browse space of teh OPC server.  I suggest that you only use
' the MoveDown and MoveToRoot functions to traverse the browse
' space.  You will see more in the frmAddItem form code.
'
Function MoveUp()
'Set error handling for OPC Function
    On Error GoTo ShowOPCMoveUpError
    
    OPCBrowserObject.MoveUp
    MoveUp = True
    GoTo SkipOPCMoveUpError
    
ShowOPCMoveUpError:
    Call DisplayOPC_COM_ErrorValue("MoveUp", Err.Number)
    MoveUp = False
SkipOPCMoveUpError:

End Function

' The GetItemCount function servers two purposes depending the
' function calls that proceed it.  If the ShowBranches function
' is called prior to GetItemCount, the value returned here will
' be the number of branches available under the current browse
' position of the OPC server.
' If the ShowLeafs function is called prior to GetItemCount,
' the value returned here will be the number of leafs(Items/Tags)
' available at the current browse position of the OPC server.
'
Function GetItemCount()
'Set error handling for OPC Function
    On Error GoTo ShowOPCGetItemCountError
    
    GetItemCount = OPCBrowserObject.Count
    GoTo SkipOPCGetItemCountError
    
ShowOPCGetItemCountError:
    Call DisplayOPC_COM_ErrorValue("GetItemCount", Err.Number)
    GetItemCount = 0
SkipOPCGetItemCountError:

End Function

' The ShowBranches function will cause the OPC server to load the
' the OPCBrowse object's item collection with any branches available
' below the current browse position.  You can then determine if there
' are any branches by calling the the GetItemCount function.  If there
' are branches available they can be retrieved using the GetItemName
' function.  As described above the OPCBrowse object only keeps track
' of one collection at a time.  If you call the ShowBranches function
' you should immediately call the GetItemCount function then get all of
' the branch names in the collection using the GetItemName function.
'
' If you call ShowBranches then call ShowLeafs, the OPCBrowse object
' will load it's internal collection first with the brances then
' overwrite the branch collection with the Leafs collection thus
' losing the content and count of branches.
'
Function ShowBranches()
'Set error handling for OPC Function
    On Error GoTo ShowOPCShowBranchesError
    
    OPCBrowserObject.ShowBranches
    ShowBranches = True
    GoTo SkipOPCShowBranchesError
    
ShowOPCShowBranchesError:
    Call DisplayOPC_COM_ErrorValue("ShowBranches", Err.Number)
    ShowBranches = False
SkipOPCShowBranchesError:

End Function

' The ShowLeafs function will cause the OPC server to load the
' the OPCBrowse object's item collection with any leafs(item/tags)
' available at the current browse position.  You can then determine
' if there are any leafs by calling the the GetItemCount function.
' If there are leafs available they can be retrieved using the
' GetItemName function.  As described above the OPCBrowse object only
' keeps track of one collection at a time.  If you call the ShowLeafs
' function you should immediately call the GetItemCount function then
' get all of the leaf names in the collection using the GetItemName
' function.
'
' If you call ShowLeafs then call ShowBranches, the OPCBrowse object
' will load it's internal collection first with the leafs then
' overwrite the leaf collection with the branch collection thus
' losing the content and count of leafs.
'
Function ShowLeafs(Optional ByVal Mode As Variant)
'Set error handling for OPC Function
    On Error GoTo ShowOPCShowLeafsError
    
    OPCBrowserObject.ShowLeafs Mode
    ShowLeafs = True
    GoTo SkipOPCShowLeafsError
    
ShowOPCShowLeafsError:
    Call DisplayOPC_COM_ErrorValue("ShowLeafs", Err.Number)
    ShowLeafs = False
SkipOPCShowLeafsError:

End Function

' The GetItemName function returns the name of the item in the
' OPCBrowse object's collection specified by the "ItemSpecifier".
' The OPCBrowse object's collection is loaded when you call either
' ShowBranches or ShowLeafs.  For the same reasons described in those
' functions you must retrieve all of the item names in a given
' collection before changing the collection.
' The "ItemSpecifier" should not exceed the value returned by the
' GetItemCount function.
'
Function GetItemName(ByRef ItemName As String, ByVal ItemSpecifier As Variant)
'Set error handling for OPC Function
    On Error GoTo ShowOPCGetItemNameError
    
    ItemName = OPCBrowserObject.Item(ItemSpecifier)
    GetItemName = True
    GoTo SkipOPCGetItemNameError
    
ShowOPCGetItemNameError:
    Call DisplayOPC_COM_ErrorValue("GetItemName", Err.Number)
    GetItemName = False
SkipOPCGetItemNameError:

End Function

' The three following functions allow you to filter the amount of
' data retuned in by either the ShowBranches or ShowLeafs functions.
' In this example I don't allow the branches to be filtered only
' leafs(items/tags).
' Filtering the number of leafs retuned in the OPCBroswe objects
' collection can make find a specific itme in the server far easier.
' This becomes apparent quickly in servers with large tag counts.
' Using these three function you can reduce the number of tags
' by filtering the tag's name, Datatype, or Access method.

' The SetFilter function allows you to specify a filter to be
' used on the acutal name of a leaf(item/tag).  By the default there
' is no filter set. This can be specified by either calling this
' funciton with an empty string "", or by calling this function with
' a string of "*".
'
' To reduce the number of items returned in the ShowLeafs function you
' can use wildcard characters such as *, ?, and name fragments such as
' "Position_*".  Using this example, the SowLeafs function would return
' all items that started with the string "Position_".
'
Function SetFilter(ByVal Filter As String)
'Set error handling for OPC Function
    On Error GoTo ShowOPCSetFilterError
    
    OPCBrowserObject.Filter = Filter
    SetFilter = True
    GoTo SkipOPCSetFilterError
    
ShowOPCSetFilterError:
    Call DisplayOPC_COM_ErrorValue("SetFilter", Err.Number)
    SetFilter = False
SkipOPCSetFilterError:
   
End Function

' In addition to filtering on the name of an item you can also filter
' based on the data type of an item.  If you want to see only itme that
' are defined as a data type of 16 bit signed integer (VT_I2) you can
' use this function set the filter data type to short (VT_I2).
'
Function SetDataTypeFilter(ByVal DataType As Integer)
'Set error handling for OPC Function
    On Error GoTo ShowOPCSetDataTypeFilterError
    
    OPCBrowserObject.DataType = DataType
    SetDataTypeFilter = True
    GoTo SkipOPCSetDataTypeFilterError
    
ShowOPCSetDataTypeFilterError:
    Call DisplayOPC_COM_ErrorValue("SetDataTypeFilter", Err.Number)
    SetDataTypeFilter = False
SkipOPCSetDataTypeFilterError:
   
End Function

' The SetAccessFilter function like the SetDataTypeFilter function
' allows you to filter the number of items retuned based on the Access
' rights of the item.  The Access rights in this cases are the item's
' Read or Write access.
'
Function SetAccessFilter(ByVal Access As Long)
'Set error handling for OPC Function
    On Error GoTo ShowOPCSetAccessFilterError
    
    OPCBrowserObject.AccessRights = Access
    SetAccessFilter = True
    GoTo SkipOPCSetAccessFilterError
    
ShowOPCSetAccessFilterError:
    Call DisplayOPC_COM_ErrorValue("SetAccessFilter", Err.Number)
    SetAccessFilter = False
SkipOPCSetAccessFilterError:
   
End Function

' Once you have browsed to a particular branch and possibly filter the
' list of items down to the few items of interest you still need to
' get the desired item's ItemID string.  The GetItemID function allows
' you to get the item ID of the "Leaf" parameter.  The leaf passed
' to this function must be available from the current browse
' position of the OPC Server.
'
' In many cases the GetItemID will return a string that closely
' matches the branch organization of the OPC server being browsed.
' You should not count on this however.  The format of the branches
' and leafs presented by the OPC server via it's browse space
' may be soley for benifit of the human user.  The ItemID returned
' may not match this branch.leaf format.
'
Function GetItemID(ByVal Leaf As String)
'Set error handling for OPC Function
    On Error GoTo ShowOPCGetItemIDError
    
    GetItemID = OPCBrowserObject.GetItemID(Leaf)
    GoTo SkipOPCGetItemIDError
    
ShowOPCGetItemIDError:
    Call DisplayOPC_COM_ErrorValue("GetItemID", Err.Number)
    GetItemID = ""
SkipOPCGetItemIDError:

End Function


' Handles displaying any OPC/COM/VB errors that are caught by the exception handler
Sub DisplayOPC_COM_ErrorValue(OPC_Function As String, ErrorCode As Long)
    Dim Response
    Dim ErrorDisplay As String
    ErrorDisplay = "The OPC function '" + OPC_Function + "' has returned an error of " + Str(ErrorCode) + " or Hex 0x" + Hex(ErrorCode)
    Response = MsgBox(ErrorDisplay, vbOKOnly, "OPC Function Error")
End Sub

