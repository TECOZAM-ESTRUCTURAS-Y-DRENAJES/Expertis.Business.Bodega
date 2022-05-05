Public Class BdgMunicipio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgMunicipio"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDMunicipio As String, ByVal services As ServiceProvider)
        Dim ClsMun As New BdgMunicipio
        If ClsMun.SelOnPrimaryKey(IDMunicipio).Rows.Count = 0 Then
            ApplicationService.GenerateError("El Municipio: | no existe en la tabla: |.", IDMunicipio, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDMunicipio As String, ByVal Services As ServiceProvider)
        Dim ClsMun As New BdgMunicipio
        If ClsMun.SelOnPrimaryKey(IDMunicipio).Rows.Count > 0 Then
            ApplicationService.GenerateError("El Municipio: | ya existe en la tabla: |.", IDMunicipio, cnEntidad)
        End If
    End Sub

#End Region

End Class