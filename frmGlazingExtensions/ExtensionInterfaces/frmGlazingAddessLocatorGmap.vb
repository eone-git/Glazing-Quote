Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Linq
Imports System.Windows.Forms
Imports DevExpress.XtraEditors
Imports GMap.NET.MapProviders
Imports GMap.NET.WindowsForms
Imports GMap.NET.WindowsForms.Markers
Imports GMap.NET
Imports System.Net
Imports Newtonsoft.Json

Partial Public Class frmGlazingAddessLocatorGmap

    Private latitude As Double = 0

    Private longitude As Double = 0

    Private newAddress As List(Of Result) = Nothing

    Private setLocation As Boolean = False

    Public Sub New(ByVal Lat As Double, ByVal Lng As Double)
        InitializeComponent()
        addressMap.Manager.Mode = GMap.NET.AccessMode.ServerAndCache
        addressMap.MapProvider = GMapProviders.GoogleMap
        addressMap.MinZoom = 2
        addressMap.MaxZoom = 24
        addressMap.Zoom = 2
        addressMap.CanDragMap = True
        addressMap.DragButton = System.Windows.Forms.MouseButtons.Left
        Me.latitude = Lat
        Me.longitude = Lng
        
    End Sub

    Private Sub BookingLocation_Load(ByVal sender As Object, ByVal e As EventArgs)
        Me.InitializeMapPonint(Me.latitude, Me.longitude)
    End Sub

    Private Sub addressMap_MouseDoubleClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles addressMap.MouseDoubleClick
        Try
            Dim lat As Double = addressMap.FromLocalToLatLng(e.X, e.Y).Lat
            Dim lng As Double = addressMap.FromLocalToLatLng(e.X, e.Y).Lng
            Me.InitializeMapPonint(lat, lng)
            Me.CreateAddress(lat, lng)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnSelectLocation_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub InitializeMapPonint(ByVal Lat As Double, ByVal Lng As Double)
        Try
            Dim overlay As GMapOverlay = New GMapOverlay("markers")
            Dim marker As GMarkerGoogle = New GMarkerGoogle(New PointLatLng(Lat, Lng), GMarkerGoogleType.red)
            marker.ToolTipMode = MarkerTooltipMode.OnMouseOver
            addressMap.Position = New PointLatLng(Lat, Lng)
            addressMap.Overlays.Clear()
            overlay.Markers.Add(marker)
            addressMap.Overlays.Add(overlay)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])

        End Try
    End Sub

    Public Function ReturnNewAddress() As List(Of Result)
        Dim address As List(Of Result) = Nothing
        address = Me.newAddress
        Return address
    End Function

    Public Sub GetAddress(ByVal addtessText As String)
        Try
            Dim addressString = GetAddressString(addtessText)
            Dim client As WebClient = New WebClient()
            AddHandler client.DownloadStringCompleted, AddressOf client_DownloadStringCompleted
            Dim Url As String = "https://maps.googleapis.com/maps/api/geocode/json?address=" & addressString & "&key=AIzaSyBvLeDvT5cURTlZUdemH99aEMTjSgffyf8"
            client.DownloadStringAsync(New Uri(Url, UriKind.RelativeOrAbsolute))
            setLocation = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try

    End Sub

    Function GetAddressString(ByVal addtessText As String) As String
        Dim addressString
        addressString = addtessText.Replace(" "c, "+"c)
        addressString = addressString.Replace(","c, "+"c)
        Return addressString

    End Function

    Sub SetMapLocation()
        latitude = newAddress.Item(0).geometry.location.lat
        longitude = newAddress.Item(0).geometry.location.lng
        addressMap.Position = New PointLatLng(latitude, longitude)
        addressMap.Zoom = 18
        setLocation = False
        Me.InitializeMapPonint(latitude, longitude)
    End Sub

    Private Sub CreateAddress(ByVal Lat As Double, ByVal Lng As Double)
        Try
            Dim client As WebClient = New WebClient()
            AddHandler client.DownloadStringCompleted, AddressOf client_DownloadStringCompleted
            Dim Url As String = "https://maps.googleapis.com/maps/api/geocode/json?address=" & Lat & "," & Lng & "&key=AIzaSyBvLeDvT5cURTlZUdemH99aEMTjSgffyf8"

            client.DownloadStringAsync(New Uri(Url, UriKind.RelativeOrAbsolute))
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub client_DownloadStringCompleted(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        Try
            Dim root = JsonConvert.DeserializeObject(Of RootObject)(e.Result)
            If root.status = "OK" Then
                newAddress = root.results
                lblAddress.Text = newAddress.Item(0).formatted_address
                If setLocation = True Then
                    SetMapLocation()

                End If

            Else
                lblAddress.Text = "No mathing address for selected Location"

            End If

        Catch ex As Exception
            MessageBox.Show("Opps.! " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.Close()

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Me.GetAddress(utxtAddress.Text)
        'SetMapLocation()
    End Sub

End Class
