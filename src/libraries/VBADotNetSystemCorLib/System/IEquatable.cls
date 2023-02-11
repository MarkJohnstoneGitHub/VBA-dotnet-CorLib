VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEquatable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Defines a generalized method that a value type or class implements to create a type-specific method for determining equality of instances."
'Rubberduck annotations
'@Folder "VBADotNetCorLib.System"
'@ModuleDescription "Defines a generalized method that a value type or class implements to create a type-specific method for determining equality of instances."
'@Interface

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 11, 2023
'@LastModified February 11, 2023

'@DotNetReferences
' https://learn.microsoft.com/en-us/dotnet/api/system.iequatable-1?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/IEquatable.cs

''
'@Remarks
'   This interface is implemented by types whose values can be equated (for example, the numeric
'   and string classes). A value type or class implements the Equals method to create a type-specific
'   method suitable for determining equality of instances.
'
'@Note
'   The IComparable<T> interface defines the CompareTo method, which determines the sort order of
'   instances of the implementing type. The IEquatable<T> interface defines the Equals method,
'   which determines the equality of instances of the implementing type.
'
'   The IEquatable<T> interface is used by generic collection objects such as Dictionary<TKey,TValue>,
'   List<T>, and LinkedList<T> when testing for equality in such methods as Contains, IndexOf,
'   LastIndexOf, and Remove. It should be implemented for any object that might be stored in a
'   generic collection.

'@Notes to Implementers
'   Replace the type parameter of the IEquatable<T> interface with the type that is implementing
'   this interface.
'
'   If you implement IEquatable<T>, you should also override the base class implementations of
'   Equals(Object) and GetHashCode() so that their behavior is consistent with that of the Equals(T) method.
'   If you do override Equals(Object), your overridden implementation is also called in calls to the
'   static Equals(System.Object, System.Object) method on your class. In addition, you should
'   overload the op_Equality and op_Inequality operators. This ensures that all tests for equality
'   return consistent results.
'
'   For information on overriding Equals(Object), see Equals(Object).
'
'   For a value type, you should always implement IEquatable<T> and override Equals(Object) for
'   better performance. Equals(Object) boxes value types and relies on reflection to compare two
'   values for equality. Both your implementation of Equals(T) and your override of Equals(Object)
'   should return consistent results.
'
'   If you implement IEquatable<T>, you should also implement IComparable<T> if instances of your
'   type can be ordered or sorted. If your type implements IComparable<T>, you almost always also
'   implement IEquatable<T>.
'
'   Note that there are some designs where a type supports an order relation, but equality may be
'   distinct from an ordering relation. Consider a Person class where you sort alphabetically.
'   Two people with the same name sort the same, but are not the same person.
''

Option Explicit

''
'@Description "Indicates whether the current object is equal to another object of the same type."
'@Parameters
'   other Variant
'       An object to compare with this object.
'
'@Returns Boolean
'   true if the current object is equal to the other parameter; otherwise, false.
'
'@Remarks
'   The implementation of the Equals method is intended to perform a test for equality with another
'   object of type T, the same type as the current object. The Equals(T) method is called in the
'   following circumstances:
'
'   When the Equals method is called and the other argument is a strongly-typed object of type T.
'   (If other is not of type T, the base Object.Equals(Object) method is called. Of the two methods,
'   IEquatable<T>.Equals offers slightly better performance.)
'
'   When the search methods of a number of generic collection objects are called. Some of these types
'   and their methods include the following:
'
'       Some of the generic overloads of the BinarySearch method.
'
'       The search methods of the List<T> class, including List<T>.Contains(T), List<T>.IndexOf,
'       List<T>.LastIndexOf, and List<T>.Remove.
'
'       The search methods of the Dictionary<TKey,TValue> class, including ContainsKey and Remove.
'
'       The search methods of the generic LinkedList<T> class, including LinkedList<T>.Contains and Remove.
'
'   In other words, to handle the possibility that objects of a class will be stored in an array or
'   a generic collection object, it is a good idea to implement IEquatable<T> so that the object can
'   be easily identified and manipulated.
'
'   When implementing the Equals method, define equality appropriately for the type specified by the
'   generic type argument. For example, if the type argument is Int32, define equality appropriately
'   for the comparison of two 32-bit signed integers.
''
Public Function Equals(ByRef other As Variant) As Boolean
'public bool Equals (T? other);
End Function
