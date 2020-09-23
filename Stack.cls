VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================================
'==                    Stack Class                    ==
'==                  By James Brannan                 ==
'==                   (C) 17 Feb 03                   ==
'=======================================================
'==                                                   ==
'== A class to simulate the Stack (Heap)              ==
'== General information:                              ==
'== When you POP a value, if no parameter is passed,  ==
'== the last item on the stack is removed. A Pushed   ==
'== value will be appended to the bottom of the stack.==
'== The first item in the stack holds the number of   ==
'== items currently on the stack. If three items were ==
'== pushed, then array offset 0 will hold 2 (0, 1, 2  ==
'== makes three items), offset 1 will hold the first  ==
'== item. The Address of an item on the stack +1 is   ==
'== the array's offset. TheStack[1] would have an     ==
'== address of 0. An example of how to use this class:==
'==                                                   ==
'=======================================================

'Dim x As New Stack
'Set x = New Stack
'x.Push ("Item #1")
'x.Push ("Item #2")
'x.Push ("Item #3")
'x.Pop
'x.Pop x.GetAdd("Item #2")
'Set x = Nothing

'Three items are added to the stack, then Item #3 is
'popped off, followed by Item #2.

Private Type StackProperty 'User-defined stack type
Value As String
Address As Integer
End Type

Public ErrUnderrun As Boolean
'Whether or not to error if you try to pop an empty stack
Public StkLimit As Integer
'The upper boundary of stack
Private Register() As StackProperty 'Disallow direct access to stack

Private Sub Class_Initialize()
ReDim Preserve Register(0) 'Create the array
Register(0).Value = -1
StkLimit = -1 'Set stack limit to unlimited
End Sub

Function Push(Val As String) As Integer 'Add value to stack, return address
If UBound(Register) + 1 > StkLimit And StkLimit > -1 Then
'If the stack has reached it's limit, raise an exception
    Call Err.Raise(772, "Push Procedure", "Buffer overrun (Stack limit has been reached)")
End If

ReDim Preserve Register(UBound(Register) + 1) 'Reallocate register
With Register(UBound(Register))
.Address = UBound(Register) - 1 'Set address of added item
.Value = Val 'Set new value of item
Push = .Address 'Return address of additem item
End With

Register(0).Value = UBound(Register) - 1
'Set register base as the number of items on the stack
End Function

Function Pop(Optional Address As Integer = 0) As String
'Remove item, return the items value
Dim Temp() As StackProperty
'Create a temporary stack to process the altered list
Dim Counter As Integer
'(Address+1 is the array offset)

If UBound(Register) = 0 And ErrUnderrun = True Then
'If we are at the bottom of the stack, and errors are allowed, display error
    Call Err.Raise(771, "Pop Procedure", "Cannot pop item from empty stack.")
End If

If Address = 0 Then Address = Register(UBound(Register)).Address
'If an address isn't provided, pop last item on stack

ReDim Temp(UBound(Register) - 1)
'Initialise temporary array with correct size
'(Size of register minus 1 popped item)
Counter = 0

For x = 0 To UBound(Register) 'Go through each item in stack
    If Not Register(x).Address = Address Then 'If item isn't the one to be popped,
    Temp(Counter).Value = Register(x).Value 'Add it to the temporary stack
    Temp(Counter).Address = Counter 'As well as it's address
    Counter = Counter + 1
    Else 'If the current item is the one to be popped,
    Pop = Register(x).Value  'Set the function to return the popped item
    End If
Next
Counter = 0 'Reset counter

ReDim Register(UBound(Temp))
'Reinitialise the real stack with one less item
Register = Temp
'Set the stack to the temporary stack
Register(0).Value = UBound(Register) - 1
'Set stack base to the number of items on stack
End Function

Function DebugStack() As Variant
Dim tmp As String
'Routine to return the items on stack, delimited by ASCII 230 (æ)

For x = 0 To UBound(Register) 'Loop through each item on stack
tmp = tmp & Register(x).Value & Chr(230) 'Add to temporary variable
Next

DebugStack = Split(tmp, Chr(230)) 'Purge the items and return them
tmp = "" 'Reset temporary variable, just in case
End Function

Function StackCount() As Integer
StackCount = Register(0).Value 'Retrieve base value on stack
End Function

Function GetVal(Address As Integer) As String
'Return value of the item on stack with the address [Address]
For x = 0 To UBound(Register) 'Loop through the stack
    If Register(x).Address = Address Then 'Check if the item's address matches
        GetVal = Register(x).Value 'If so, return it
        Exit Function 'And exit the loop
    End If
Next
End Function

Function GetAdd(Value As String) As Integer
'Returns the address of the item on stack matching the value
'Note that if more than one item has the value [Value], the first one is returned
For x = 0 To UBound(Register) 'Loop through the stack
    If Register(x).Value = Value Then 'If the value matches,
        GetAdd = Register(x).Address 'Return it
        Exit Function 'And exit the loop
    End If
Next
End Function
