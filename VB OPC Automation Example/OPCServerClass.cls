VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OPCServerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The OPCServerClass object is the basic building block object
' for your OPC application.  All interaction with an OPC server
' begins first by connecting to the OPC server.  The OPCServerClass
' object wraps the Automation interface OPCServer object.
' This diagram shows the relationships between the various components
' of the Automation interface wrapper.
'
' >  OPCServer                                      <
' >          |                                      <
' >          +-OPC Groups (Collection)              <
'                      |
'                      +-OPC Group
'                                |
'                                +-OPC Items (Collection)
'                                           |
'                                           +-OPC Items
'
' The OPCServerClass object handles functions found in the first two
' branches of this diagram, the OPCServer object and OPCGroups collection.
' I have not implemented all of the methods and properties found in these
' objects.  Adding access to any methods or properties not currently in
' this module is easy.
'
' The OPCServerClass object handles encapsulating the ServerShutDown
' event.  This allows each instance of an the OPCServerClass object
' you create to get its own ServerShutDown event.

Option Explicit
Option Base 1
' The OPCServer objects must be declared here due to the use of WithEvents
' The OPCServer object is the basic object of the the Automation
' Interface.
Dim WithEvents OPCServerObj As OPCServer
Attribute OPCServerObj.VB_VarHelpID = -1

' The server name is the Prog ID of the connected OPC Server. This
' string is passed to the ConnectOPCServer function of this module
' from the VB application.
Dim OPCServerName As String

' This is a string key used to reference this OPCServerClass object.  The
' OPCServerKey can be used by the VB application anyway the user
' sees fit.  In some cases this is used as a key in controls such as
' a treeview of a listview.
Dim OPCServerKey As String

' The ServerGroups object contains the OPCGroups collection of the
' Automation Interface.  This is used to add new groups to the OPC
' server.  This is shown in the function AddOPCGroup.  This collection
' is kept by the actual Automation Interface wrapper.  The OPCServerClass
' object also keeps its own collection of OPCGroupClass objects.
Dim ServerGroups As OPCGroups

' The OPCServerGroups collection keeps list of OPC groups just like the
' ServerGroups object above.  This collection however is a list of
' OPCGroupClass objects.  The OPCGroupClass object wrap the methods and
' properties of the Automation Interface's OPCGroup and OPCItems collection.
Dim OPCServerGroups As New Collection

' The OPCServerIndex is the numeric version of the OPCServerClass key.
Dim OPCServerIndex As Integer

' The OPCServerClass object has the ability to generate an event when
' an OPC server is shutting down.   When an OPC server supports the
' OPC 2.0 data access specification, it has the option of issuing a
' ServerShutDown event to any of it's attached clients.  In the case
' of a VB application the OPCServer oject will signal the
' ServerShutDown event.  This event is handled here in the OPCServerClass
' object.  If this event fires, the ServerShuttingDown event of the
' OPCServerClass object will be fired.  A VB application has the option of
' hooking this event to be informed when the OPC connection is going to be
' lost.
Public Event ServerShuttingDown(ByVal ServerKey As String)


' The GetOPCServerList function is used to get a list of OPC servers
' available on either the local machine or a remote machine by
' supplying an optional node name.  The GetOPCServerList function
' can be used without an existing OPC connection as you would expect
' otherwise we would have chicken and egg problem.
'
Sub GetOPCServerList(ByRef ServerList As Variant, Optional ByVal NodeName As Variant)
    Dim i As Integer
    'Set error handling for OPC Function
    On Error GoTo ShowOPCGetServersError
    
'    Dim myNodeName As Variant
'    myNodeName = "192.168.57.137"
'    myNodeName = "yy-2k"
'    ServerList = OPCServerObj.GetOPCServers(myNodeName)
    ServerList = OPCServerObj.GetOPCServers(NodeName)
     
    GoTo SkipOPCGetServersError
    
ShowOPCGetServersError:
    Call DisplayOPC_COM_ErrorValue("Get OPC Server List", Err.Number)
SkipOPCGetServersError:
End Sub


' This sub handles connecting with the selected OPC Server
' The OPCServer Object provides a method called 'Connect' that
' allows you to 'connect' with an OPC server.  The 'Connect'
' method can take two arguments, a server name and a Node name.
' In this example we don't use the Node name since it is
' optional.  When the 'Connect' method is called you should see
' the OPC Server application start if it is not aleady running.
' This function requires that the ServerKey be passed in as string
' forcing the VB application to manage the server keys.  Once a
' OPCServerClass object has been created and connected, the OPCGroupClass
' and OPCItemClass object create their own keys.
'
Function ConnectOPCServer(ServerName As String, ServerKey As String, ServerIndex As Integer, Optional ByVal NodeName As Variant)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCConnectError
    '
    'Create a new OPC Server object
    'Load the selected server name to start the interface
    Dim StoreName As String
    Dim StoreKey As String
    
    StoreName = ServerName
    OPCServerName = StoreName
    OPCServerIndex = ServerIndex
    
    ' The ServerKey is stored as part of the OPCServerClass object to
    ' make access the collections and treeview easier.
    StoreKey = ServerKey
    OPCServerKey = StoreKey
    
    'Attempt to connect with the server
    '*****************************************************************
    '                                                                *
    '                  (Local only in this example)                  *
    '                                                                *
    '*****************************************************************
    OPCServerObj.Connect OPCServerName, NodeName

    ' Establish the base group collection object to add new groups
    ' this object will be used below in the group interface functions
    ' At thismpoint there are no groups we simply have the interface
    ' that will allow groups to be added.
    Set ServerGroups = OPCServerObj.OPCGroups
    
    ' Establish the initial default conditions for new groups added to this
    ' server.
    SetDefaultGroupIsActive (True)
    SetDefaultGroupUpdateRate (100)
    SetDefaultGroupDeadBand (0)
    
    
    ConnectOPCServer = True
    GoTo SkipOPCConnectError
    
ShowOPCConnectError:
    Call DisplayOPC_COM_ErrorValue("Connect", Err.Number)
    ConnectOPCServer = False
SkipOPCConnectError:
End Function


' This sub handles disconnecting from the OPC Server.  The OPCServer
' Object provides the method 'Disconnect'.  Calling this on an
' active OPCServer object will release the OPC Server interface with
' your application.  When this occurs you should see the OPC server
' application shut down if it started automatically on the OPC
' connect. This function should not be called until the group(s) and items
' have been removed from this OPCServerClass object.
'
Function DisconnectOPCServer()
    'Set error handling for OPC Function
    On Error GoTo ShowOPCDisconnectError

    'Disconnect from the server, This should only occur after the
    'items and group have been removed.  The server will release
    ' the groups and items for this connection even if you don't
    ' but it is good practice to release the groups and items
    ' progmatically since you don't want to count on the OPC
    ' server to clean up after your application.
    OPCServerObj.Disconnect
    DisconnectOPCServer = True
    
    GoTo SkipDisconnectError
ShowOPCDisconnectError:
    Call DisplayOPC_COM_ErrorValue("Disconnect", Err.Number)
    DisconnectOPCServer = False
SkipDisconnectError:
End Function

' This function handles returning the OPCServerKey string. The
' OPCServerKey is local to the OPCSererClass object and does not
' interact with the underlying Automation Interface.
'
Function GetOPCServerKey()
    GetOPCServerKey = OPCServerKey
End Function

' This functions returns a reference to the OPCServerGroups
' collection maintained by the OPCServerClass object. The
' OPCServerGroups is local to the OPCSererClass object and is not
' retrieved from the underlying Automation Interface.
'
Function GetOPCServerGroupCollection()
    Set GetOPCServerGroupCollection = OPCServerGroups
End Function


' This function returns Serverindex as a numeric value.  The
' OPCServerIndex is local to the OPCSererClass object and does not
' interact with the underlying Automation Interface.
'
Function GetOPCServerIndex()
    GetOPCServerIndex = OPCServerIndex
End Function


' This function returns the name of the OPC Server, this is the ProgId
' of the OPC server.
'
Function GetServerName(ByRef ServerName As String)
    ' Attempt to get the server ServerName.
    ServerName = OPCServerName
    GetServerName = True
End Function


' The next set of functions provide wrappers to the OPCServer object
' methods that allow you to get the various properties of the OPCServer
' object.
'


' Get the server StartTime property
' The server start time returns the time the server has started on
' the host machine.  The server start time is unique to the
' server. All connected clients will receive the same time.
'
Function GetStartTime(ByRef StartTime As Date)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCStartTimeError
    
    ' Attempt to get the server start time.
    StartTime = OPCServerObj.StartTime
    GetStartTime = True
    GoTo SkipOPCStartTimeError
    
ShowOPCStartTimeError:
    Call DisplayOPC_COM_ErrorValue("StartTime", Err.Number)
    GetStartTime = False
    
SkipOPCStartTimeError:
End Function


' Get the server CurrentTime property
' The value is the current time at the server.  It is acquired
' by the OPC automation interface using the GetStatus interface.
'
Function GetCurrentTime(ByRef CurrentTime As Date)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCCurrentTimeError
    
    ' Attempt to get the server current time.
    CurrentTime = OPCServerObj.CurrentTime
    GetCurrentTime = True
    GoTo SkipOPCCurrentTimeError
    
ShowOPCCurrentTimeError:
    Call DisplayOPC_COM_ErrorValue("CurrentTime", Err.Number)
    GetCurrentTime = False
    
SkipOPCCurrentTimeError:
End Function


' Get the LastUpdateTime property
' The LastUpdateTime return the time the last update occured
' for the server this value is for all clients connected to
' the server therfore the time returned may not be the last
' time your client received an update.
'
Function GetLastUpdateTime(ByRef LastUpdateTime As Date)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCLastUpdateTimeError
    
    ' Attempt to get the server Last Update time.
    LastUpdateTime = OPCServerObj.LastUpdateTime
    GetLastUpdateTime = True
    GoTo SkipOPCLastUpdateTimeError
    
ShowOPCLastUpdateTimeError:
    Call DisplayOPC_COM_ErrorValue("LastUpdate", Err.Number)
    GetLastUpdateTime = False
    
SkipOPCLastUpdateTimeError:
End Function


' Get the major version number of the connected server
'
Function GetMajorVersion(ByRef MajorVersion As Integer)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCMajorVersionError
    
    ' Attempt to get the server MajorVersion.
    MajorVersion = OPCServerObj.MajorVersion
    GetMajorVersion = True
    GoTo SkipOPCMajorVersionError
    
ShowOPCMajorVersionError:
    Call DisplayOPC_COM_ErrorValue("MajorVersion", Err.Number)
    GetMajorVersion = False
    
SkipOPCMajorVersionError:
End Function


' Get the minor version number of the connected server
'
Function GetMinorVersion(ByRef MinorVersion As Integer)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCMinorVersionError
    
    ' Attempt to get the server MinorVersion.
    MinorVersion = OPCServerObj.MinorVersion
    GetMinorVersion = True
    GoTo SkipOPCMinorVersionError
    
ShowOPCMinorVersionError:
    Call DisplayOPC_COM_ErrorValue("MinorVersion", Err.Number)
    GetMinorVersion = False
    
SkipOPCMinorVersionError:
End Function


' Get the build number of the connected server
'
Function GetBuildNumber(ByRef BuildNumber As Integer)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCBuildNumberError
    
    ' Attempt to get the server BuildNumber.
    BuildNumber = OPCServerObj.BuildNumber
    GetBuildNumber = True
    GoTo SkipOPCBuildNumberError
    
ShowOPCBuildNumberError:
    Call DisplayOPC_COM_ErrorValue("BuildNumber", Err.Number)
    GetBuildNumber = False
    
SkipOPCBuildNumberError:
End Function


' Get the Vendor Info of the connected server
' The vendor info will return a string describing the
' name of the vendor of the OPC server.  In some cases this string
' will return the name of the OPC Server toolkit vendor if
' the connected OPC server uses a toolkit.
'
Function GetVendorInfo(ByRef VendorInfo As String)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCVendorInfoError
        
    ' Attempt to get the server VendorInfo.
    VendorInfo = OPCServerObj.VendorInfo
    GetVendorInfo = True
    GoTo SkipOPCVendorInfoError
    
ShowOPCVendorInfoError:
    Call DisplayOPC_COM_ErrorValue("VendorInfo", Err.Number)
    GetVendorInfo = False
    
SkipOPCVendorInfoError:
End Function


' Get the Server State property of the connected server
' The value retured will be one of the following
' OPC_STATUS_RUNNING      - 1 - Server is running normally
' OPC_STATUS_FAILED       - 2 - Vendor specific fatal error has occured
' OPC_STATUS_NOCONFIG     - 3 - Server Running but no Configuration Data available
' OPC_STATUS_SUSPENDED    - 4 - Server is suspended and not receiving data
' OPC_STATUS_TEST         - 5 - Server in test mode
' OPC_STATUS_DISCONNECTED - 6 - Server has disconnected
'
Function GetServerState(ByRef ServerState As Long)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCServerStateError
    
    ' Attempt to get the server ServerState.
    ServerState = OPCServerObj.ServerState
    GetServerState = True
    GoTo SkipOPCServerStateError
    
ShowOPCServerStateError:
    Call DisplayOPC_COM_ErrorValue("ServerState", Err.Number)
    GetServerState = False
    
SkipOPCServerStateError:
End Function

' This function returns an OPC browse object from the server object
' held in this instance of the OPCServerClass.  This function may return
' a true condition indicating that the function succeeded, however the
' the OPCBrowserObject may still be set to Nothing if the connected
' server does not support browsing.  With that in mind check the
' returned object for "Nothing" before use.

Function GetServerBrowseObject(ByRef OPCBrowserObject As OPCBrowserClass)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCBrowserError
    
    Dim OPCBrowserObj As OPCBrowser
        
    ' Attempt to get the server browse object
    Set OPCBrowserObj = OPCServerObj.CreateBrowser
    If Not OPCBrowserObj Is Nothing Then
        Set OPCBrowserObject = New OPCBrowserClass
        OPCBrowserObject.SetBrowserObject OPCBrowserObj
    Else
        Set OPCBrowserObject = Nothing
    End If
    GetServerBrowseObject = True
    GoTo SkipOPCBrowserError
    
ShowOPCBrowserError:
    Call DisplayOPC_COM_ErrorValue("Browser Object", Err.Number)
    GetServerBrowseObject = False
    
SkipOPCBrowserError:
End Function

' The next section of this module covers the management of OPC groups
' within the OPCServerClass object. When adding an OPC group the initial
' state of the groups atrributes must be set.  A group has the following
' attributes, GroupName, Update Rate, DeadBand, and ActiveState.  The
' actual ServerGroups.Add function only allows you to specify the group
' name when adding a new group.  To set the attributes listed above you
' can use the Set Default functions that follow to establish the initial
' conditions for a new group.



' The SetDefaultGroupIsActive allows you to set the default active
' condition on the ServerGroups object determining the active
' state of a newly added group.  Once a group is added you can change
' these attributes by using the functions found in the OPCGroupClass
' object. The active state of a group allows you to literally turn the
' data in this group on or off.  For the OPC server this means that
' your application either needs the data in the group(On/Active) or
' it doesn't(Off/Inactive).  Using the active state of the group
' your VB application can tailor how much data the OPC server is trying
' read from your devices.  By turing a group off that you don't
' currently need you can reduce the amount of data the OPC server is
' reading from a device.  This of course has the benefit of increasing
' the performance of the OPC server for the data items that you do still
' need to acquire.  I recommend using the active state of the group or
' of the item(shown later), instead of continuously adding and removing
' items from a group as a means of controlling what data the OPC server
' is gathering from the devices.
'
Function SetDefaultGroupIsActive(ByVal State As Boolean)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCDefaultGroupIsActiveError
    
    ServerGroups.DefaultGroupIsActive = State
    SetDefaultGroupIsActive = True
    GoTo SkipOPCDefaultGroupIsActiveError
    
ShowOPCDefaultGroupIsActiveError:
    Call DisplayOPC_COM_ErrorValue("DefaultGroupIsActive", Err.Number)
    SetDefaultGroupIsActive = False
    
SkipOPCDefaultGroupIsActiveError:
End Function


' The SetDefaultGroupUpdateRate allows you to set the default UpdateRate
' condition on the ServerGroups object determining the Update Rate
' of a newly added group.  Once a group is added you can change
' these attributes by using the functions found in the OPCGroupClass
' object. The update rate is entered in milliseconds.  The range is
' 0 to 2147483647 milliseconds which is about 596 hours. Entering
' value of 0 tells the OPC server you want it to acquire data in this
' group as fast as possible.  The update rate on a group is very
' powerful tool for optimizing your application's communications.
' If you set every group you add to a server to an update interval of
' 0 the server will attempt to gather all of your data as quickly as
' possible.  If all of your data requires the best possible update
' speed then this is probably OK.  But what if some of your data doesn't
' really change but once every 10 minutes or so.  To read this as fast
' as possible would waste a considerable amount of the communications
' bandwidth of your OPC server.  Since the OPC server allows you to add
' multiple groups to the server you can separate your data into separate
' groups, with each group having its own update rate.  In this way fast
' data can be gathered quickly and slow data can be gathered slowly.
'
Function SetDefaultGroupUpdateRate(ByVal Rate As Long)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCDefaultUpdateRateError
    
    ServerGroups.DefaultGroupUpdateRate = Rate
    SetDefaultGroupUpdateRate = True
    GoTo SkipOPCDefaultUpdateRateError
    
ShowOPCDefaultUpdateRateError:
    Call DisplayOPC_COM_ErrorValue("SetDefaultGroupUpdateRate", Err.Number)
    SetDefaultGroupUpdateRate = False
    
SkipOPCDefaultUpdateRateError:
End Function


' The SetDefaultGroupDeadBand allows you to set the default Deadband
' condition on the ServerGroups object determining the DeadBand
' of a newly added group.  Once a group is added you can change
' these attributes by using the functions found in the OPCGroupClass
' object.  Deadband is an interesting feature of the an OPC server.
' Deadband allows you to limit the amount of data changes the OPC
' will return for items in this group.  It does this by establishing the
' ratio of change that must occur in a given item before the OPC server
' will pass the changed value back to the connected OPC client.
' The Deadband ranges from 0 to 100 percent. The actual percentage of
' effect that the deadband has upon changing data depends of the
' data tpye of the OPC item.  If the data type of the OPC item is a BYTE
' which has a range of 0 to 255 a deadband value of 1 percent would
' require that the value contained in this BYTE item must change by
' atleast 2.5 before the OPC server will inform the OPC client of the
' change.  Unlike the the Active state and Update Rate of a group the
' Deadband feature will not optimize the communications bandwidth of the
' OPC server.  Deadband is more likely to be used to prevent an
' application like a data logger from logging every little change in
' value due to a noisy enviroment as an example.
'
Function SetDefaultGroupDeadBand(ByVal DeadBand As Single)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCDefaultGroupDeadBandError
    
    ServerGroups.DefaultGroupDeadband = DeadBand
    SetDefaultGroupDeadBand = True
    GoTo SkipOPCDefaultGroupDeadBandError
    
ShowOPCDefaultGroupDeadBandError:
    Call DisplayOPC_COM_ErrorValue("SetDefaultGroupDeadBand", Err.Number)
    SetDefaultGroupDeadBand = False
    
SkipOPCDefaultGroupDeadBandError:
End Function


' The name say it all, gets the default active state from the
' ServerGroups collection object.
'
Function GetDefaultGroupIsActive(ByRef State As Boolean)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCGetDefaultGroupIsActiveError
    
    State = ServerGroups.DefaultGroupIsActive
    GetDefaultGroupIsActive = True
    GoTo SkipOPCGetDefaultGroupIsActiveError
    
ShowOPCGetDefaultGroupIsActiveError:
    Call DisplayOPC_COM_ErrorValue("GetDefaultGroupIsActive", Err.Number)
    GetDefaultGroupIsActive = False
    
SkipOPCGetDefaultGroupIsActiveError:
End Function


' The name say it all, gets the default UpdateRate from the
' ServerGroups collection object.
'
Function GetDefaultGroupUpdateRate(ByRef Rate As Long)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCGetDefaultUpdateRateError
    
    Rate = ServerGroups.DefaultGroupUpdateRate
    GetDefaultGroupUpdateRate = True
    GoTo SkipOPCGetDefaultUpdateRateError
    
ShowOPCGetDefaultUpdateRateError:
    Call DisplayOPC_COM_ErrorValue("GetDefaultGroupUpdateRate", Err.Number)
    GetDefaultGroupUpdateRate = False
    
SkipOPCGetDefaultUpdateRateError:
End Function


' The name say it all, gets the default DeadBand from the
' ServerGroups collection object.
'
Function GetDefaultGroupDeadBand(ByRef DeadBand As Single)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCGetDefaultGroupDeadBandError
    
    DeadBand = ServerGroups.DefaultGroupDeadband
    GetDefaultGroupDeadBand = True
    GoTo SkipOPCGetDefaultGroupDeadBandError
    
ShowOPCGetDefaultGroupDeadBandError:
    Call DisplayOPC_COM_ErrorValue("GetDefaultGroupDeadBand", Err.Number)
    GetDefaultGroupDeadBand = False
    
SkipOPCGetDefaultGroupDeadBandError:
End Function


' This sub handles adding the group to the OPC server and
' establishing the group interface.  When adding a group you
' can preset some of the group parameters using the properties
' '.DefaultGroupIsActive' and '.DefaultGroupDeadband'.  Set these
' before adding the group. Once the group has been successfully
' added you can change these same settings along with the group
' update rate on the fly using the properties on the resulting
' OPCGroupClass object.
' The GroupName parameter is the name your would like to use for
' the new group to be added.  This paramter can be left blank.
' When you leave the group name blank you are telling the OPC server
' that is should generate a group name for us.  See the SetDefault
' functions above for details on the UpdateRate, DeadBand, and active
' state of the group.
'
' The GroupKey is string that will be returned to the VB application.
' This string can be used as a key for this group in a collection.
'
Function AddOPCGroup(GroupName As String, UpdateRate As Long, DeadBand As Single, ActiveState As Boolean, ByRef GroupKey As String)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCGroupAddError
    
    ' Create an new instance of the OPCGroupClass
    Dim ConnectedGroup As New OPCGroupClass
    Dim NewGroup As OPCGroup
    Dim GroupNum As Integer
              
    ' Establish the initial default conditions for new groups added to this
    ' server.
    SetDefaultGroupIsActive (ActiveState)
    SetDefaultGroupUpdateRate (UpdateRate)
    SetDefaultGroupDeadBand (DeadBand)

    ' Normally you can name an OPC group yourself but
    ' you also have the option of allowing the OPC Server
    ' to provide you with a unique group name automatically.
    ' To do this simply leave the GroupName empty.
    ' This blank group name is then passed to the OPC
    ' server which will recognize this as a automatic
    ' group name.  The server will return the name it
    ' generates as part of the Automation Interface's
    ' OPCGroup object.
    
    Set NewGroup = ServerGroups.Add(GroupName)
    
    
    ' If the GroupName was passed in as blank then get the group
    ' name that the OPC Server has automatically generated.
    If GroupName = "" Then
        GroupName = NewGroup.Name
    End If
    
    ' The FindNextGroupNumber handles getting a unique group number
    ' if you continue to add groups to your OPCServerClass object
    ' this function simply returns the next group number.  If you
    ' delete a group that was previously added this function will
    ' attempt to find the deleted group numbers and use them first.
    ' This prevents the group index numbers from just continuing to
    ' increment.
    GroupNum = FindNextGroupNumber
    
    ' Once the next group number is found a unique group key is
    ' developed.  This key can be used in your application to manage
    ' access to this group.  One example might be to use the GroupKey
    ' returned here as a key within a TreeView of ListView control.
    ' As you can see here the group key is comprise of the group
    ' index specifically for this group and on the server index.
    ' By combining these two index numbers in one key we can be
    ' resonably sure that the group key will be unique across
    ' multiple OPC server groups.
    GroupKey = "Group" + Str(GroupNum) + Str(OPCServerIndex)
    
    ' On the outside chance the OPC server didn't generated a group name
    ' we can create one from the Groupkey.
    If GroupName = "" Then
        GroupName = GroupKey
    End If
   
    ' This sets the basic properties of the OPCGroupClass object
    ConnectedGroup.SetOPCGroup NewGroup, GroupName, GroupKey, GroupNum
    
    ' Although the OPCGroups (ServerGroups) object is already a
    ' collection we need to have our own OPCGGroupClass object wrapper
    ' to contain the properties and events for each group we intend
    ' to add to the server.  This means that we must keep the
    ' ServerGroups collection and the OPCServerGroups collection
    ' in sync.
    With OPCServerGroups
        .Add ConnectedGroup, GroupKey
    End With
        
    AddOPCGroup = True
    GoTo SkipAddGroupError

ShowOPCGroupAddError:
    Call DisplayOPC_COM_ErrorValue("Add Group", Err.Number)
    AddOPCGroup = False
    
SkipAddGroupError:
End Function


' This function attempts to keep the Group index number
' from continuosly growing by finding a group number that
' may have been deleted and returning those first before
' going to the next group index number. The end result of
' this function it that the group number will never grow
' beyond actual number of groups.
'
Private Function FindNextGroupNumber()
    ' In this function an error state means that a vacant
    ' group number has been found.
    On Error GoTo FoundNextGroupNumber
    
    Dim i As Integer
    Dim GroupNum As Integer
    
    GroupNum = 1
    
    With OPCServerGroups
        ' This loop tries to access each item in the OPCServerGroups
        ' collection by it's Key.  If an item can not be pulled from the
        ' collection that means the group had been deleted and it's
        ' group number is available.
        For i = 1 To .Count
            GroupNum = i
            .Item ("Group" + Str(i) + Str(OPCServerIndex))
        Next i
        
        ' If we get here all of the group numbers in the collection
        ' exist and a next available group number should be returned.
        If .Count <> 0 Then
            FindNextGroupNumber = i
        Else
            FindNextGroupNumber = 1 ' No count return the first 1
        End If
    End With
    
    GoTo NewGroup
    
FoundNextGroupNumber:
    ' If we get here a group number could not be found in the
    ' OPCServerGroups collection which means that it has been deleted
    ' and is available for a new group addition.
    FindNextGroupNumber = GroupNum
    
NewGroup:
End Function


' This sub handles removing a group from the OPC server.  The group
' is removed from the ServerGroups object which is the collection
' managed by the Automation Interface.  Then the group is removed from
' the OPCServerGroups collection which is managed by the OPCServerClass
'
Function RemoveOPCGroup(GroupKey As String)
    'Set error handling for OPC Function
    On Error GoTo ShowOPCGroupRemoveError
      
    Dim OPCGroupCls As OPCGroupClass
    Set OPCGroupCls = OPCServerGroups.Item(GroupKey)
    Dim GroupName As String
    GroupName = OPCGroupCls.GetGroupName
      
    ' First the group is removed from the actual group collection managed by
    ' the automation interface if it is successful then remove the tiem from the
    ' OPCServerGroups collection
    ServerGroups.Remove GroupName
    
    ' Now we remove it from the OPCServerGroups collection
    ' to allow any final processing on the group to be done
    OPCServerGroups.Remove GroupKey
        
    RemoveOPCGroup = True
    GoTo SkipRemoveGroupError

ShowOPCGroupRemoveError:
    Call DisplayOPC_COM_ErrorValue("Remove Group", Err.Number)
    RemoveOPCGroup = False
    
SkipRemoveGroupError:
End Function


' The OPCServerClass object has the ability to generate an event when
' an OPC server is shutting down.   When an OPC server supports the
' OPC 2.0 data access specification, it has the option of issuing a
' ServerShutDown event to any of it's attached clients.  In the case
' of a VB application the OPCServer oject will signal the
' ServerShutDown event.  This event is handled here in the OPCServerClass
' object.  If this event fires, the ServerShuttingDown event of the
' OPCServerClass object will be fired.  A VB application has the option of
' hooking the ServerShuttingDown event to be informed when the OPC
' connection is going to be lost.
'
Private Sub OPCServerObj_ServerShutDown(ByVal Reason As String)
    ' If we receive a server shut down message we need to release everthing and remove the server
    ' connection.
    RaiseEvent ServerShuttingDown(OPCServerKey)
End Sub


' Create the new OPCServer object when the OPCServerClass object is
' instantiated.
'
Private Sub Class_Initialize()
    Set OPCServerObj = New OPCServer
End Sub


' When the OPCServerClass object is deleted this function will cleanup
' the OPC connection by releasing the groups and items.  While this function
' will clean up the groups and items that may have been attached to this
' OPCServerClass object you will still need to call DisconnectOPCServer
' in your VB application on this object to properly release the OPC Server
' from the Automation Interface.
'
Private Sub Class_Terminate()
    
    ' Normally you should remove all groups and their items from the server
    ' before you remove the server object.  However the OPCGroups collection
    ' object supports a RemoveAll method that will remove all items from
    ' the groups and all groups from the server.
    If Not ServerGroups Is Nothing Then
        If ServerGroups.Count <> 0 Then
            ServerGroups.RemoveAll
        End If
    End If
    
    ' We still need to remove the OPCServerGroups collection that contains
    ' the OPCGroupClasss objects
    If OPCServerGroups.Count <> 0 Then
        Dim i As Integer
        Dim a As Integer
        a = OPCServerGroups.Count
        For i = 1 To OPCServerGroups.Count
            With OPCServerGroups
            .Remove (a)
            a = a - 1
            End With
        Next i
    End If
    
    ' Release the OPCServerGroups collection
    Set OPCServerGroups = Nothing
    
End Sub


' Handles displaying any OPC/COM/VB errors that are caught by the exception handler
Sub DisplayOPC_COM_ErrorValue(OPC_Function As String, ErrorCode As Long)
    Dim Response
    Dim ErrorDisplay As String
    ErrorDisplay = "The OPC function '" + OPC_Function + "' has returned an error of " + Str(ErrorCode) + " or Hex 0x" + Hex(ErrorCode)
    Response = MsgBox(ErrorDisplay, vbOKOnly, "OPC Function Error")
End Sub

