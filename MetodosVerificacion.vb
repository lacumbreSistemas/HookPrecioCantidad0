Imports QSRules
Public Class MetodosVerificacion
    Public Session As QSRules.SessionClass
    Public Correcto As Boolean = True

    Public Sub RutinaVerificacion()
        For index As Integer = 1 To Session.Transaction.Entries.Count
            Try
                If Session.Transaction.Entries(index).Price <= 0 Then
                    MsgBox("El producto '" + Session.Transaction.Entries(index).Item.Description + "' tiene precio en 0 o menor y no se puede vender, Llamar a un supervisor.")
                    Correcto = False
                ElseIf Session.Transaction.Entries(index).Quantity <= 0 Then
                    MsgBox("El producto  '" + Session.Transaction.Entries(index).Item.Description + "'  tiene Cantidad en 0 o menor y no se puede vender, Llamar a un supervisor.")
                    Correcto = False
                ElseIf Session.Transaction.Entries(index).Item.Inactive = True Then
                    MsgBox("El producto  '" + Session.Transaction.Entries(index).Item.Description + "'  esta Inactivo y no se puede vender, Llamar a un supervisor.")
                    Correcto = False
                End If
            Catch ex As Exception
                Correcto = False
            End Try
        Next
    End Sub
End Class
