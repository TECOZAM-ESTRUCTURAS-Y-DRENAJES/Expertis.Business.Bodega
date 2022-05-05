Public Class BdgProveedor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgProveedor"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDProveedor As String) As DataRow
        Dim dt As DataTable = New BdgProveedor().SelOnPrimaryKey(IDProveedor)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el proveedor | en la gestión de bodega.", IDProveedor)
        Else : Return dt.Rows(0)
        End If
    End Function

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DataGetTarifa
        Public IDProveedor As String = String.Empty
        Public IDVariedad As String = String.Empty
        Public Vendimia As Integer

        Public Sub New(ByVal IDProveedor As String, ByVal IDVariedad As String, ByVal Vendimia As Integer)
            Me.IDProveedor = IDProveedor
            Me.IDVariedad = IDVariedad
            Me.Vendimia = Vendimia
        End Sub
    End Class

    <Task()> Public Shared Function GetTarifa(ByVal data As DataGetTarifa, ByVal services As ServiceProvider) As Tarifa
        Dim r As New Tarifa

        '1º Buscamos la tarifa de la Proveedor/Variedad
        Dim dtProveedorVariedad As DataTable = New BdgProveedorVariedad().SelOnPrimaryKey(data.IDProveedor, data.IDVariedad)
        If Not dtProveedorVariedad Is Nothing AndAlso dtProveedorVariedad.Rows.Count > 0 Then
            Dim rwProveedorVariedad As DataRow = dtProveedorVariedad.Rows(0)
            r.IDTarifaT = rwProveedorVariedad(_P.IDTarifaT) & String.Empty
            r.PrecioOrigenT = rwProveedorVariedad(_P.PrecioOrigenT)
            r.PrecioExcedenteT = rwProveedorVariedad(_P.PrecioExcedenteT)
            r.IDTarifaB = rwProveedorVariedad(_P.IDTarifaB) & String.Empty
            r.PrecioOrigenB = rwProveedorVariedad(_P.PrecioOrigenB)
            r.PrecioExcedenteB = rwProveedorVariedad(_P.PrecioExcedenteB)
        End If

        '2º Buscamos la tarifa del Proveedor
        Dim rwProveedor As DataRow = New BdgProveedor().GetItemRow(data.IDProveedor)
        r.IDTarifaT = rwProveedor(_P.IDTarifaT) & String.Empty
        r.PrecioOrigenT = rwProveedor(_P.PrecioOrigenT)
        r.PrecioExcedenteT = rwProveedor(_P.PrecioExcedenteT)
        r.IDTarifaB = rwProveedor(_P.IDTarifaB) & String.Empty
        r.PrecioOrigenB = rwProveedor(_P.PrecioOrigenB)
        r.PrecioExcedenteB = rwProveedor(_P.PrecioExcedenteB)

        '3º Buscamos la tarifa del Grupo del Proveedor
        If Not r.Completo Then
            Dim strIDGrupo As String
            If Length(rwProveedor(_P.IDGrupo)) > 0 Then strIDGrupo = rwProveedor(_P.IDGrupo)
            If Len(strIDGrupo) Then
                Dim rwGrupo As DataRow = New BdgGrupo().GetItemRow(strIDGrupo)

                If Length(r.IDTarifaT) = 0 And Length(rwGrupo(_G.IDTarifaT)) > 0 Then r.IDTarifaT = rwGrupo(_G.IDTarifaT)
                If r.PrecioOrigenT = 0 Then r.PrecioOrigenT = rwGrupo(_G.PrecioOrigenT)
                If r.PrecioExcedenteT = 0 Then r.PrecioExcedenteT = rwGrupo(_G.PrecioExcedenteT)
                If Length(r.IDTarifaB) = 0 And Length(rwGrupo(_G.IDTarifaB)) > 0 Then r.IDTarifaB = rwGrupo(_G.IDTarifaB)
                If r.PrecioOrigenB = 0 Then r.PrecioOrigenB = rwGrupo(_G.PrecioOrigenB)
                If r.PrecioExcedenteB = 0 Then r.PrecioExcedenteB = rwGrupo(_G.PrecioExcedenteB)
            End If
        End If

        '4º Buscamos la tarifa de la Variedad
        If Not r.Completo Then
            If Len(data.IDVariedad) > 0 Then
                Dim rwVariedad As DataRow = New BdgVariedad().GetItemRow(data.IDVariedad)

                If Length(r.IDTarifaT) = 0 And Length(rwVariedad(_VDM.IDTarifaT)) > 0 Then r.IDTarifaT = rwVariedad(_VDM.IDTarifaT)
                If r.PrecioOrigenT = 0 Then r.PrecioOrigenT = rwVariedad(_VDM.PrecioOrigenT)
                If r.PrecioExcedenteT = 0 Then r.PrecioExcedenteT = rwVariedad(_VDM.PrecioExcedenteT)
                If Length(r.IDTarifaB) = 0 And Length(rwVariedad(_VDM.IDTarifaB)) > 0 Then r.IDTarifaB = rwVariedad(_VDM.IDTarifaB)
                If r.PrecioOrigenB = 0 Then r.PrecioOrigenB = rwVariedad(_VDM.PrecioOrigenB)
                If r.PrecioExcedenteB = 0 Then r.PrecioExcedenteB = rwVariedad(_VDM.PrecioExcedenteB)
            End If
        End If

        '5º Buscamos la tarifa de la Vendimia
        If Not r.Completo Then
            If data.Vendimia Then
                Dim rwVendimia As DataRow = New BdgVendimia().GetItemRow(data.Vendimia)

                If Length(r.IDTarifaT) = 0 And Length(rwVendimia(_VDM.IDTarifaT)) > 0 Then r.IDTarifaT = rwVendimia(_VDM.IDTarifaT)
                If r.PrecioOrigenT = 0 Then r.PrecioOrigenT = rwVendimia(_VDM.PrecioOrigenT)
                If r.PrecioExcedenteT = 0 Then r.PrecioExcedenteT = rwVendimia(_VDM.PrecioExcedenteT)
                If Length(r.IDTarifaB) = 0 And Length(rwVendimia(_VDM.IDTarifaB)) > 0 Then r.IDTarifaB = rwVendimia(_VDM.IDTarifaB)
                If r.PrecioOrigenB = 0 Then r.PrecioOrigenB = rwVendimia(_VDM.PrecioOrigenB)
                If r.PrecioExcedenteB = 0 Then r.PrecioExcedenteB = rwVendimia(_VDM.PrecioExcedenteB)
            End If
        End If

        Return r

    End Function

#End Region

End Class

<Serializable()> _
Public Class _P
    Public Const IDProveedor As String = "IDProveedor"
    Public Const IDGrupo As String = "IDGrupo"
    Public Const IDTarifaT As String = "IDTarifaT"
    Public Const IDTarifaB As String = "IDTarifaB"
    Public Const PrecioOrigenT As String = "PrecioOrigenT"
    Public Const PrecioOrigenB As String = "PrecioOrigenB"
    Public Const PrecioExcedenteT As String = "PrecioExcedenteT"
    Public Const PrecioExcedenteB As String = "PrecioExcedenteB"
End Class

<Serializable()> _
Public Class Tarifa
    Private mIDTarifaT As String
    Private mIDTarifaB As String
    Private mPrecioOrigenT As Double
    Private mPrecioOrigenB As Double
    Private mPrecioExcedenteT As Double
    Private mPrecioExcedenteB As Double

    Public Function Completo() As Boolean
        Return Len(IDTarifaT) > 0 AndAlso Len(IDTarifaB) > 0 AndAlso PrecioOrigenT <> 0 AndAlso PrecioOrigenB <> 0 AndAlso mPrecioExcedenteT <> 0 AndAlso mPrecioExcedenteB <> 0
    End Function

    Public Property IDTarifaT() As String
        Get
            Return mIDTarifaT
        End Get
        Set(ByVal Value As String)
            If Len(mIDTarifaT) = 0 Then mIDTarifaT = Value
        End Set
    End Property

    Public Property IDTarifaB() As String
        Get
            Return mIDTarifaB
        End Get
        Set(ByVal Value As String)
            If Len(mIDTarifaB) = 0 Then mIDTarifaB = Value
        End Set
    End Property

    Public Property PrecioOrigenT() As Double
        Get
            Return mPrecioOrigenT
        End Get
        Set(ByVal Value As Double)
            If mPrecioOrigenT = 0 Then mPrecioOrigenT = Value
        End Set
    End Property

    Public Property PrecioOrigenB() As Double
        Get
            Return mPrecioOrigenB
        End Get
        Set(ByVal Value As Double)
            If mPrecioOrigenB = 0 Then mPrecioOrigenB = Value
        End Set
    End Property

    Public Property PrecioExcedenteT() As Double
        Get
            Return mPrecioExcedenteT
        End Get
        Set(ByVal Value As Double)
            If mPrecioExcedenteT = 0 Then mPrecioExcedenteT = Value
        End Set
    End Property

    Public Property PrecioExcedenteB() As Double
        Get
            Return mPrecioExcedenteB
        End Get
        Set(ByVal Value As Double)
            If mPrecioExcedenteB = 0 Then mPrecioExcedenteB = Value
        End Set
    End Property
End Class