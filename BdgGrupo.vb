Public Class BdgGrupo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgGrupo"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDGrupo As String) As DataRow
        Dim Dt As DataTable = New BdgGrupo().SelOnPrimaryKey(IDGrupo)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el grupo |", IDGrupo)
        Else : Return dt.Rows(0)
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _G
    Public Const IDGrupo As String = "IDGrupo"
    Public Const DescGrupo As String = "DescGrupo"
    Public Const IDTarifaT As String = "IDTarifaT"
    Public Const IDTarifaB As String = "IDTarifaB"
    Public Const PrecioOrigenT As String = "PrecioOrigenT"
    Public Const PrecioOrigenB As String = "PrecioOrigenB"
    Public Const PrecioExcedenteT As String = "PrecioExcedenteT"
    Public Const PrecioExcedenteB As String = "PrecioExcedenteB"
End Class