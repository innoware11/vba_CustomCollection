# vba_CustomCollection
This class represents a class that allows to work with the VBA Collection by keys. 
Class adds methods:
  - Add(key As String, value As Variant) - Method for adding an element to a collection with a given key
  - Item(key As String) As Variant - Method for getting value by key
  - Remove(key As String) - Method for deleting an element by key
  - Exists(key As String) As Boolean - Method for checking the presence of an element by key
  - Keys() As Collection - Method to get all keys of a collection
  - Items() as Collection - Method to get all values of a collection

Usage - add to Class Modules CustomCollection.
Dim collection as new CustomCollection
use class methods    
  
