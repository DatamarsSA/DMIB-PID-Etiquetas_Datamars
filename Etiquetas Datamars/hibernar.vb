Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class hibernar
    'Clase que sirve para desactivar la hibernación de los equipos, así como el bloqueo de pantalla.
    Public Sub New()

    End Sub
    'importamos una libreria de win32
    <DllImport("KERNEL32.DLL", SetLastError:=True)>
    Public Shared Function SetThreadExecutionState(ByVal esFlags As executionState
       ) As executionState
    End Function
    'Creamos una Enumeracion con los valores que debemos pasarle a libreria importada
    Public Enum executionState As UInteger
        None = &H0UI 'ningun valor se usa para comprobar si hay errores en la funcion.
        SystemRequired = &H1UI 'reinicia el temporizador de inactividad del sistema
        DisplayRequired = &H2UI 'reinicia el temporizador de inactividad de la pantalla.
        UserPresent = &H4UI 'modelo de usuario presente. No se usa da fallo.
        Continuous = &H80000000UI 'Informa al sistema el estado que debe de tener los flags del sistema.
        AwaymodeRequired = &H40UI 'Activa el modo ausente, solo usarlo en aplicaciones multimedia
    End Enum
    Public Sub deshibernacion()
        'Procedimiento para desactivar la hibernación.
        Dim result As executionState = executionState.None
        'Recibimos el valor dado por la función. Borramos los flags y le damos los nuevos valores, reiniciamos el sistema, reiniciamos la pantalla.
        result = SetThreadExecutionState(executionState.Continuous Or executionState.SystemRequired Or executionState.DisplayRequired)
        If result = executionState.None Then MsgBox("Se ha producido un error al Deshabilitar la hibernación del equipo")

    End Sub
    Public Sub actHibernacion()
        'Procedimiento para activar la hibernación.
        Dim result As executionState = executionState.None
        'Recibimos el valor dado por la función. Borramos los flags con los valores anteriores y coge lo que tiene el sistema por defecto.
        result = SetThreadExecutionState(executionState.Continuous)
        If result = executionState.None Then MsgBox("Se ha producido un error al habilitar la hibernación del equipo")

    End Sub

End Class
