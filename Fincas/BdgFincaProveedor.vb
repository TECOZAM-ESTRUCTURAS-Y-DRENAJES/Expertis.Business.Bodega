Public Class BdgFincaProveedor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFincaProveedor"

#End Region

End Class

<Serializable()> _
Public Class _FP
    Public Const IDFinca As String = "IDFinca"
    Public Const IdProveedor As String = "IDProveedor"
    Public Const Porcentaje As String = "Porcentaje"
End Class