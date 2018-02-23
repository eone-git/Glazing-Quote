
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Public Class AddressComponent

    Public Property long_name As String

    Public Property short_name As String

    Public Property types As List(Of String)
End Class

Public Class Bounds

    Public Property northeast As Location

    Public Property southwest As Location
End Class

Public Class Location

    Public Property lat As Double

    Public Property lng As Double
End Class

Public Class Geometry

    Public Property bounds As Bounds

    Public Property location As Location

    Public Property location_type As String

    Public Property viewport As Bounds
End Class

Public Class Result

    Public Property address_components As List(Of AddressComponent)

    Public Property formatted_address As String

    Public Property geometry As Geometry

    Public Property partial_match As Boolean

    Public Property types As List(Of String)
End Class

Public Class RootObject

    Public Property results As List(Of Result)

    Public Property status As String
End Class

