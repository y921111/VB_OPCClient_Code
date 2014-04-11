Attribute VB_Name = "Module1"
' The Automation interface provides the Visual Basic user with a powerful
' set of objects to access OPC servers.  The Automation Wrapper for OPC
' allows you to use every feature available in the OPC specification.
' The Data Access Automation Interface Standard goes a long way towards
' describing the various aspects of each automation object's methods,
' properties, or events but it still leaves a lot to be desired when it
' comes to providing a good real world example of how you can use this
' powerful new tool in your applications.
'
' The goal of this example is provide a simple yet full featured project
' that you can use as a starting point to produce your own applications
' that can take advantage of many facilities available to you when using
' OPC servers as your source of communications connectivity.  While it is
' hoped that you use our own KEPServerEX product, this example will work
' with any OPC server you wish to use.  That is the whole point of OPC
' in the first place, giving you the user the widest range of options
' for solving your communications needs.  If you are using KEPServerEX
' load the supplied sample application called "Simdemo".  This application
' will allow you to browse and add preconfigured OPC tags that have
' changing data.
'
'
' Let me start by saying that I don't claim to be a Visual basic
' programmer so there may be a couple of things that the more seasoned
' VB programmer would do differently and that's fine by me.  My goal
' was to make the example as complete as possible operationally as well
' as robust.  All OPC functions are housed in 4 class modules
' OPCServerClass, OPCGroupClass, OPCItemClass and OPCBrowserClass.  Each
' of these class modules depend upon each other and are structured to
' reflect the object oriented relationship between the OPC Automation
' interface and it's own objects.  The remainder of code is found in
' the various Form modules used to handle the user interface.  Don't
' discount these modules, a lot of little tidbits of code is housed in
' these modules. The main form, frmMain is primary display module, all
' user interface actions start there.  This module also houses much of
' the code that demonstrates how the 3 class modules are used.
'
' One of the things I found right away is that while the Automation
' Wrapper goes a long way it doesn't provide enough of a frame work to
' build a truely dynamic application.  Many of our customers liked my
' previous simple VB example but it only allowed 1 server connection
' 1 group, and 10 little items all in a very fixed set of static
' variables and arrays. While it could be expanded to handle more
' OPC items easily it was still really limited to use with very simple
' applications.  With this new VB example I wanted to provide a framework
' for a truley dynamic OPC client in VB.  This example will allow you to
' connect to any number of OPC servers, add any number of OPC groups to
' those servers and any number of OPC items to those groups.  I hope
' that you find this example to be benficial to you.
'
' I recommend reading the "Data Access Automation Interface Standard"
' before diving into this example.  As I said above the spec
' may not give all the details one would hope but it will help you
' understand the nature of the objects provided by the Automation Wrapper
' and how I have used them here in this example.
'
' Now for that ever popular legal stuff.
'
' This programming example is provided "AS IS", as such KEPware, Inc.
' makes no claims to the worthiness of the code and does not warranty
' the code to be error free.  It is provided freely and can be used in
' your own projects.  If you do find this code useful place a little
' marketing plug for KEPware in your code.  While I would love to be
' able to support everyone who calls about these code examples the very
' nature of VB makes each and every one of your projects unique, as such it
' will be difficult to support every application, besides this is what
' I do from 9 pm. to 4 am.  If you really find yourself in a bind please
' contact KEPware's technical support. If your using an OPC server
' other than KEPServer or KEPServerEX we won't be able to help on
' the server side if that's where you problems or questions center.
' With that said here we go.



'**********************************************************************
'                             Code Starts Here

' These are two OPC constants that are used when accessing OPC item
' data.  I wanted these to be global to the entire project so I placed
' them here normally they would have been in a header file for one of
' the class modules but we don't have that in VB land so here they are.

Public Const OPCItemDirect = 1
Public Const OPCItemLocal = 2


' The fMainForm variable is used to provide global access to
' the operating instance of the main form module.
Public fMainForm As frmMain

' This varible is used to generate a unigue key that will be
' used in the OPCServers collection as a string based key
' The key is "Server" + Str(ServerIndex) resulting in "Server 1"
' , "Server 2" and so on.  This is a little short cut for the
' server connections, the groups and items handle there own index
' keying internally and dynamically adjust for additions and deletions.
Public ServerIndex As Integer

' The OPCServers collection is the basic starting point for all of your
' OPC connections.  This collection will contain a list of OPCServerClass
' objects as you add server connections to your application.  It is this
' collection that uses the string key developed and described above in the
' ServerIndex.
Public OPCServers As New Collection

' This next set of variables are used to contain a working set
' of pointers to the objects the user has selected on this
' application's interface.  These variables are used across all of
' the various user interface forms to determine which OPC object the
' user intends to add, change or invoke.
'
' The SelectedOPCServer variable will be set when the user has added
' an OPC server connection and then either Left or Right clicks on
' the server connection in the tvTreeView on the main form.
Public SelectedOPCServer As OPCServerClass

' The SelectedOPCGroup variable will be set when the user has added
' an OPC group to an existing server connection and then either
' Left or Right clicks on the group in the tvTreeView on the main form.
' In this example you can count on the SelectedOPCServer also being set
' when SelectedOPCGroup is set and valid.
Public SelectedOPCGroup As OPCGroupClass

' The SelectedOPCItem variable will be set when the user has added
' an OPC group to an existing server connection, then added an
' item to that group and either Left or Right clicks on the item in
' the lvListView on the main form.
' In this example you can count on the SelectedOPCServer and
' SelectedOPCGroup also being set when SelectedOPCItem is set and valid.
Public SelectedOPCItem As OPCItemClass



' Here we begin.  I figure you'll comment out the splash screen on
' about the second run.
Sub Main()
    ' Comment out the next two lines to ditch the splash screen.
    ' The Splash does show how to access the SetWindowPos function in
    ' the Windows API to keep it on the top while it is displayed.
    frmSplash.Show
    frmSplash.Refresh
    
    ' Create a working instance of the main form, load and display it.
    Set fMainForm = New frmMain
    Load fMainForm
    fMainForm.Show
    ' This index is used to build a unique server name for the
    ' server key in the tvTreeView control so initialize it here.
    ServerIndex = 0
End Sub

