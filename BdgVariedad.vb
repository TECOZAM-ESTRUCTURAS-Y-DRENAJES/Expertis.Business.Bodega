Public Class BdgVariedad

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVariedad"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDVariedad As String, ByVal services As ServiceProvider)
        If New BdgVariedad().SelOnPrimaryKey(IDVariedad).Rows.Count = 0 Then
            ApplicationService.GenerateError("La Variedad: | no existe en la tabla: |.", IDVariedad, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDVariedad As String, ByVal services As ServiceProvider)
        If New BdgVariedad().SelOnPrimaryKey(IDVariedad).Rows.Count > 0 Then
            ApplicationService.GenerateError("La Variedad: | ya existe en la tabla: |.", IDVariedad, cnEntidad)
        End If
    End Sub

#End Region

End Class

Public Enum BdgTipoVariedad
    Tinta = 0
    Blanca = 1
End Enum
