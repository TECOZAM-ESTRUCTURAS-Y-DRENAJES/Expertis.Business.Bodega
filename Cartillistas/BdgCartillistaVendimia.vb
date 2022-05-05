Public Class BdgCartillistaVendimia

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCartillistaVendimia"

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StDuplicarCartillista
        Public IDCartillista As String
        Public Vendimia As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCartillista As String, ByVal Vendimia As Integer)
            Me.IDCartillista = IDCartillista
            Me.Vendimia = Vendimia
        End Sub
    End Class

    <Task()> Public Shared Function DuplicarCartillista(ByVal strIDCartillista As String, ByVal lngVendimia As Integer) As DataTable
        Dim Services As New ServiceProvider
        If Length(strIDCartillista) > 0 Then
            Dim IntUltimaVendimia As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf BdgVendimia.UltimaVendimia, New Object, Services)
            If IntUltimaVendimia > 0 Then
                Dim dtCtaVendimia As DataTable = New BdgCartillistaVendimia().SelOnPrimaryKey(strIDCartillista, lngVendimia)
                If dtCtaVendimia.Rows.Count > 0 Then
                    Dim oRw As DataRow = dtCtaVendimia.Rows(0)
                    Dim rwAux As DataRow = dtCtaVendimia.NewRow
                    dtCtaVendimia.Rows.Add(rwAux)
                    If Not rwAux Is Nothing Then
                        rwAux(_CV.IDCartillista) = strIDCartillista
                        rwAux(_CV.Vendimia) = IntUltimaVendimia
                        rwAux(_CV.HaT) = oRw(_CV.HaT)
                        rwAux(_CV.HaB) = oRw(_CV.HaT)
                        Dim ClsBdgCartillista As New BdgCartillistaVendimia
                        ClsBdgCartillista.Update(rwAux.Table)
                        Return rwAux.Table
                    End If
                End If
            End If
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _CV
    Public Const IDCartillista As String = "IDCartillista"
    Public Const Vendimia As String = "Vendimia"
    Public Const NCartilla As String = "NCartilla"
    Public Const HaT As String = "HaT"
    Public Const HaB As String = "HaB"
    Public Const MaxT As String = "MaxT"
    Public Const MaxB As String = "MaxB"
    Public Const ExtT As String = "ExtT"
    Public Const ExtB As String = "ExtB"
    Public Const Talon As String = "Talon"
End Class