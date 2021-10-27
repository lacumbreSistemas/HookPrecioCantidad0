<ComClass(Main.ClassId, Main.InterfaceId, Main.EventsId)> _
Public Class Main

#Region "GUID de COM"
    ' Estos GUID proporcionan la identidad de COM para esta clase 
    ' y las interfaces de COM. Si las cambia, los clientes 
    ' existentes no podrán obtener acceso a la clase.
    Public Const ClassId As String = "16af76d7-a811-4c7c-bb10-817ed8608092"
    Public Const InterfaceId As String = "cccd6953-88db-40fd-9d57-1d5f7c3c0c72"
    Public Const EventsId As String = "b509e06d-c110-4413-9119-7486418854ad"
#End Region

    ' Una clase COM que se puede crear debe tener Public Sub New() 
    ' sin parámetros, si no la clase no se 
    ' registrará en el registro COM y no se podrá crear a 
    ' través de CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Function Process(ByVal currSession As Object) As Boolean
        If currSession.Transaction.TransactionType = 1 And currSession.Transaction.ReturningItems = False Then
            Dim f As New MetodosVerificacion
            f.Session = currSession
            f.RutinaVerificacion()
            Return f.Correcto
        End If
        Return True
    End Function
End Class


