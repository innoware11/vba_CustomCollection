VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class representing a custom collection
' Created on 25/04/2024
' Updated on 25/04/2024
'----------------------------
'
' Author: Maksym Shevchenko
' email: innoware11@gmail.com
' github: innoware11

' Class representing a custom collection
' Created on 25/04/2024
' Updated on 25/04/2024
'----------------------------
'
' Author: Maksym Shevchenko
' email: innoware11@gmail.com
' github: innoware11

' Class representing a custom collection
Private internalCollection As New Collection
Private keysCollection As New Collection

' Method for adding an element to a collection with a given key
Public Sub Add(key As String, value As Variant)
    ' Checking if there is an element with such a key
    If Exists(key) Then
        ' If an element with such a key already exists, remove it from the collection
        Remove key
    End If
    
    ' Add the element and key to the corresponding collections
    internalCollection.Add value
    keysCollection.Add key
End Sub

' Method for getting value by key
Public Function Item(key As String) As Variant
    Dim index As Integer
    index = GetKeyIndex(key)
    If index > 0 Then Item = internalCollection(index)
End Function

' Method for deleting an element by key
Public Sub Remove(key As String)
    ' Finding an element's index by key
    Dim index As Integer
    index = GetKeyIndex(key)
    
    If index > 0 Then
        ' Removing an element from keysCollection
        keysCollection.Remove index
        ' Removing an element from internalCollection
        internalCollection.Remove index
    End If
End Sub

' Method for checking the presence of an element by key
Public Function Exists(key As String) As Boolean
    Exists = (GetKeyIndex(key) > 0)
End Function

' Method to get all keys of a collection
Public Function Keys() As Collection
    Set Keys = keysCollection
End Function

' Method to get all values of a collection
Public Function Items() As Collection
    Set Items = internalCollection
End Function

' Function to get the index of a key in keysCollection
Private Function GetKeyIndex(key As String) As Integer
    Dim index As Integer
    For index = 1 To keysCollection.count
        If keysCollection(index) = key Then
            GetKeyIndex = index
            Exit Function
        End If
    Next index
End Function
