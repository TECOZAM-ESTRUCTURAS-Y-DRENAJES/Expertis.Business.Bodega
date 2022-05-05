Public Class ProcessInfoOperacion
    Public AutoCalcularVino As Boolean
    Public GuardarPropuesta As Boolean
    Public MultiplesOperaciones As Boolean

    Public Sub New(ByVal AutoCalcularVino As Boolean, ByVal GuardarPropuesta As Boolean, ByVal MultiplesOperaciones As Boolean)
        Me.AutoCalcularVino = AutoCalcularVino
        Me.GuardarPropuesta = GuardarPropuesta
        Me.MultiplesOperaciones = MultiplesOperaciones
    End Sub
End Class
