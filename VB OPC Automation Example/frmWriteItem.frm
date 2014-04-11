VERSION 5.00
Begin VB.Form frmWriteItem 
   Caption         =   "OPC Write Item"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton ApplyValue 
      Caption         =   "Apply Value"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Value to Write"
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.TextBox DataType 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "This box simply displays the selected Data Type for reference"
         Top             =   720
         Width           =   4935
      End
      Begin VB.OptionButton BooleanOff 
         Caption         =   "Off"
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton BooleanOn 
         Caption         =   "On"
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Top             =   1800
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox ItemID 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "This box simply displays the selected ItemID for reference"
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox WriteValue 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Enter the value to be written here, for arrays separate each element by a comma"
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Data Type:"
         Height          =   255
         Left            =   195
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Boolean Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ItemID:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Numeric and String Values"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmWriteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The OPC Write Item form is used to send new data
' to a specific OPC item. The form will automatically
' set itself to accept either numeric, string,
' or boolean input depending on the data type of the
' item selected.
'
' This write operation uses the AsyncWriteOPCItem function
' of the OPCGroupClass object The OPCGroupClass wraps the
' AsyncWrite function of the Automation interface.  The
' AsyncWrite function has the benefit of returning control
' immediately to your application.  There are other mehtods
' for doing writes to OPC items but using the Async method
' is the best way to insure that your application does not
' get blocked by the write operation of the OPC server.
' When you use the write function directly on an OPC item
' or the SyncWrite function on an OPC group the operation
' is blocked until the OPC server completes the write.
' This includes the round trip time to the physical device
' and back. Depending on your device this could be a long
' wait during which your VB application would appear to be
' locked up. Using the AsyncWrite as shown here insures
' that your VB application can continue to process other
' operations.  The results of the write operation is
' return to you later in the AsyncWriteComplete which is
' handled in the OPCGroupClass object.  This event will
' place the returned error in the specific Quality field
' of the OPC item.
'
' When the Write item form is displayed it will atempt the
' identify and dsiplay the data type of the item selected.
' If you check the VB help system for VarType Const you'll
' see that some 17 data types are listed. Many of these
' datatypes do not apply to data you will receive from an
' OPC server.  There is also the chance you will get back a
' datatype which doesn't fit one of these 17 datatypes.
' When this occurs you are most likely getting back an
' Unsigned Data type.  Most of the time the Automation
' Wrapper can handle these unsigned datatypes.  I have
' however seen a problem when attempting to write to an
' array of Unsigned short values.  This is why I force
' the user to select one of the data types when adding an
' OPC item.  I do allow the Native datatype selection which
' allows the OPC server to pick what it thinks is the
' default(canonical) datatype of the item.  This is where
' an items datatype can become unsigned.

Option Explicit
' The option Base is not set here since we need an
' array that can start with a lower bound of zero

Dim DataTypes(17) As String

' The ValueChange flag is used to determine if the
' value needs to be written when the OK button is
' pressed.  The Apply button allows you to write the
' value each time it is pressed without leaving the
' OPC Write Item form.  If you change the value without
' hitting the apply button and exit the dialog with Ok
' button the value will still be sent.  If the apply
' button is hit before the Ok button the value will not
' be written on the OK.
Dim ValueChange As Boolean


Private Sub Form_Load()
    ' Get the Item ID string for reference
    ItemID.Text = Module1.SelectedOPCItem.GetItemID
        
    ' This list of data types is derived from the VarType Constants
    ' Some of the data type in the VarType Constants don't apply to
    ' normal OPC data objects and are therefore marked as invalid
    
    DataTypes(0) = "Native - Data Type is not specific set by OPC Server"
    DataTypes(1) = "Invalid Data Type for this example"
    DataTypes(2) = "Integer - 16 bit Signed "
    DataTypes(3) = "Long - 32 bit Signed "
    DataTypes(4) = "Float - 32 bit Float "
    DataTypes(5) = "Double - 64 bit Real "
    DataTypes(6) = "Invalid Data Type for this example" ' Currency
    DataTypes(7) = "Invalid Data Type for this example" ' Date
    DataTypes(8) = "String - 8 bit Signed Characters"
    DataTypes(9) = "Invalid Data Type for this example" ' Object
    DataTypes(10) = "Invalid Data Type for this example" ' Error
    DataTypes(11) = "Boolean - True or Flase"
    DataTypes(12) = "Invalid Data Type for this example" ' Variant
    DataTypes(13) = "Invalid Data Type for this example" ' Data Access
    DataTypes(14) = "Invalid Data Type for this example" ' Decimal
    DataTypes(15) = "Invalid Data Type for this example" ' Unspecified type
    DataTypes(16) = "Invalid Data Type for this example" ' Unspecified type
    DataTypes(17) = "Byte 8 bit Unsigned "
    
    ' Now that we have a list of data type descriptions
    ' here where we use them
    If Module1.SelectedOPCItem.GetItemCanonicalType < 18 Then
        ' For normal datatypes use the strin found in the array
        DataType.Text = DataTypes(Module1.SelectedOPCItem.GetItemCanonicalType)
    ElseIf (Module1.SelectedOPCItem.GetItemCanonicalType > 8192) And (Module1.SelectedOPCItem.GetItemCanonicalType < 8210) Then
        ' For arrays the data type will the value vbArray(8192) added to the normal datatype
        ' This simply removes the array type to get the normal data type then
        ' sticks the word "Array" on the end of teh datatypes text.
        DataType.Text = DataTypes(Module1.SelectedOPCItem.GetItemCanonicalType - 8192) + "Array"
    Else
        ' You may see this if the data type of the item is an unsigned data type
        ' in these cases you may not be able to do a write.  I have seen this
        ' occur when trying to write to an array of unsigned integers.
        ' Writes to a single unsigned value have worked just fine however so
        ' doesn't appear to be a clear rule of thumb here. The issue with
        ' writes to the unsigned array appears to be an issue in the Automation
        ' Wrapper itself.
        DataType.Text = "Possible Unsigned data type, Writes may not work"
    End If
    
    ' Enable the appropriate edit control for the given data type
    If Module1.SelectedOPCItem.GetItemCanonicalType = vbBoolean Then
        BooleanOn.Enabled = True
        BooleanOff.Enabled = True
        WriteValue.Enabled = False
    Else
        BooleanOn.Enabled = False
        BooleanOff.Enabled = False
        WriteValue.Enabled = True
        If Module1.SelectedOPCItem.GetItemCanonicalType <> vbString Then
            WriteValue.Text = "0"
        Else
            WriteValue.Text = "String Data"
        End If
    End If
    ValueChange = False
    
End Sub


' On each click of the Apply button write the current
' value to the OPC item.  By setting ValueChange to True
' the WriteNewValue function will write the data each time.
'
Private Sub ApplyValue_Click()
    ValueChange = True
    WriteNewValue
End Sub


' The next two function simply create a radio button effect
' between the BooleanOn and BooleanOff check boxes.
'
Private Sub BooleanOff_Click()
    BooleanOn.Value = False
    ' Make sure the OK button cna write the value if needed
    ValueChange = True
End Sub

Private Sub BooleanOn_Click()
    BooleanOff.Value = False
    ' Make sure the OK button cna write the value if needed
    ValueChange = True
End Sub


' The Deactivate event is used here as a cheap way
' of keeping the form on top of the application if
' the user clicks on the main window.  There may be
' other ways to do this but this works fairly well.
'
Private Sub Form_Deactivate()
    ' use the show memthod to force the form
    ' back to the front
    frmWriteItem.Show
End Sub


' On OK click call the WriteNewValue it will check to see if
' ValueChange flag is true.
'
Private Sub OkButton_Click()
    WriteNewValue
    ' Reenable the tree view
    fMainForm.tvTreeView.Enabled = True
    Unload Me
End Sub


' Each time the user changes the WriteValue text box
' set the ValueChange flag true so the OK button can
' cause the value to be written.
'
Private Sub WriteValue_Change()
    ValueChange = True
End Sub


' The WriteNewValue sub handles changing the value of
' a single OPC item.  Depending on the data type the
' WriteNewValue function will interpret the proper
' input control and write the data accordingly.  For
' Boolean, Numeric, and Boolean this is fairly straight
' forward.  For arrays its a little more complex.
' When the data type is not Boolean we need to see if
' the item is possibly an array. If it is not an array
' we can check for a String type and then send the new value
' as either a single numeric or a single string.
'
' For array I allow the values to be enter in a string of
' values separated by commas.  You can even enter commas
' with no values between them to skip over elements of the
' array.  To demonstrate, if you have an array of four 16
' bit integers you can enter a string in the WriteValue
' text box of 1,2,3,4 .  This will send the new values
' 1,2,3,4 to each of the four array elements.  If you enter
' the string ,,55, you will only change the value of
' element 3.
'
Private Sub WriteNewValue()
    Dim ValueToWrite As Variant
    Dim CurrentValue As Variant
    
    ' Don't write unless there has been a change
    If ValueChange = True Then
           
        If Module1.SelectedOPCItem.GetItemCanonicalType = vbBoolean Then
            ' Write the Boolean state.  Since the two check box controls'
            ' act as radio buttons we only need to look at the state
            ' of one of them.
            ValueToWrite = BooleanOn.Value
        Else
            ' Get the current item in its variant form
            CurrentValue = Module1.SelectedOPCItem.GetItemValue(OPCItemDirect)
            
            ' Check to see if it is an array
            If Not IsArray(CurrentValue) Then
                ' its not an array see if it is not a string
                If Module1.SelectedOPCItem.GetItemCanonicalType <> vbString Then
                    ' It is not a string so convert the value in the WriteValue text box
                    ValueToWrite = Val(WriteValue.Text)
                Else
                    ' It is a string so send the string as is.
                    ValueToWrite = WriteValue.Text
                End If
            Else
                ' When writing an array we first grab
                ' the CurrentValues this allows us to change
                ' only the items we want within the array.
                ValueToWrite = CurrentValue
                Dim i As Integer
                Dim Length As Integer
                Dim LastStartPos As Integer
                Dim LastEndPos As Integer
                Length = Len(WriteValue.Text)
                Dim ArrayValue As Variant
                LastStartPos = 1
                LastEndPos = 1
                For i = 0 To UBound(ValueToWrite)
                    LastEndPos = InStr(LastStartPos, WriteValue.Text, ",")
                    If LastEndPos = 0 Then
                        LastEndPos = Length + 1
                    End If
                    ' See if there is data for this element number by checking
                    ' the gap between the LastStartPos and LastEndPos.
                    If LastStartPos < LastEndPos Then
                        ' new value for element so load into the working array
                        ValueToWrite(i) = Val(Mid(WriteValue.Text, LastStartPos, LastEndPos - LastStartPos))
                    End If
                    LastStartPos = LastEndPos + 1
                Next i  ' try to get the next array element
            End If
        End If
        
        ' Now that a Variant datatype has been developed that contains
        ' the value to be written, send it. In this case we are directly
        ' calling the AsyncWriteOPCItem of the OPCGroupClasss on the
        ' selected group.
        Module1.SelectedOPCGroup.AsyncWriteOPCItem Module1.SelectedOPCItem, ValueToWrite
    End If
    ' Now that the value has been written clear the ValueChange flag.
    ValueChange = False
End Sub

