Module Module1
    Public Structure DCB
        Public DCBlength As Int32
        Public BaudRate As Int32
        Public fBitFields As Int32
        Public wReserved As Int16
        Public XonLim As Int16
        Public XoffLim As Int16
        Public ByteSize As Byte
        Public Parity As Byte
        Public StopBits As Byte
        Public XonChar As Byte
        Public XoffChar As Byte
        Public ErrorChar As Byte
        Public EofChar As Byte
        Public EvtChar As Byte
        Public wReserved1 As Int16 'Reserved; Do Not Use
    End Structure

    Public Structure COMMTIMEOUTS
        Public ReadIntervalTimeout As Int32
        Public ReadTotalTimeoutMultiplier As Int32
        Public ReadTotalTimeoutConstant As Int32
        Public WriteTotalTimeoutMultiplier As Int32
        Public WriteTotalTimeoutConstant As Int32
    End Structure

    Public Const GENERIC_READ As Int32 = &H80000000
    Public Const GENERIC_WRITE As Int32 = &H40000000





    Public Const OPEN_EXISTING As Int32 = 3
    Public Const FILE_ATTRIBUTE_NORMAL As Int32 = &H80
    Public Const NOPARITY As Int32 = 0
    Public Const ONESTOPBIT As Int32 = 0

    Public Declare Auto Function CreateFile Lib "kernel32.dll" _
       (ByVal lpFileName As String, ByVal dwDesiredAccess As Int32,
          ByVal dwShareMode As Int32, ByVal lpSecurityAttributes As IntPtr,
             ByVal dwCreationDisposition As Int32, ByVal dwFlagsAndAttributes As Int32,
                ByVal hTemplateFile As IntPtr) As IntPtr

    Public Declare Auto Function GetCommState Lib "kernel32.dll" (ByVal nCid As IntPtr,
       ByRef lpDCB As DCB) As Boolean

    Public Declare Auto Function SetCommState Lib "kernel32.dll" (ByVal nCid As IntPtr,
       ByRef lpDCB As DCB) As Boolean

    Public Declare Auto Function GetCommTimeouts Lib "kernel32.dll" (ByVal hFile As IntPtr,
       ByRef lpCommTimeouts As COMMTIMEOUTS) As Boolean

    Public Declare Auto Function SetCommTimeouts Lib "kernel32.dll" (ByVal hFile As IntPtr,
       ByRef lpCommTimeouts As COMMTIMEOUTS) As Boolean

    Public Declare Auto Function WriteFile Lib "kernel32.dll" (ByVal hFile As IntPtr,
       ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToWrite As Int32,
          ByRef lpNumberOfBytesWritten As Int32, ByVal lpOverlapped As IntPtr) As Boolean

    Public Declare Auto Function ReadFile Lib "kernel32.dll" (ByVal hFile As IntPtr,
       ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToRead As Int32,
          ByRef lpNumberOfBytesRead As Int32, ByVal lpOverlapped As IntPtr) As Boolean

    Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal hObject As IntPtr) As Boolean

    Sub escribe_puerto(ByVal mo As String, ByVal puerto As String)
        Dim hSerialPort, hParallelPort As IntPtr
        Dim Success As Boolean
        Dim MyDCB As DCB
        Dim MyCommTimeouts As COMMTIMEOUTS
        Dim BytesWritten, BytesRead As Int32
        Dim Buffer() As Byte

        ' Declare variables to use for encoding.
        Dim oEncoder As New System.Text.ASCIIEncoding
        Dim oEnc As System.Text.Encoding = oEncoder.GetEncoding(1252)

        ' Convert String to Byte().
        Buffer = oEnc.GetBytes(mo)

        Try
            ' Parallel port.
            'Console.WriteLine("Accessing the LPT1 parallel port")
            'Obtain a handle to the LPT1 parallel port.
            hParallelPort = CreateFile("LPT1", GENERIC_READ Or GENERIC_WRITE, 0, IntPtr.Zero,
               OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)
            'Verify that the obtained handle is valid.
            If hParallelPort.ToInt32 = -1 Then
                ' Throw New CommException("Unable to obtain a handle to the LPT1 port")
            End If
            ' Retrieve the current control settings.
            '  Success = GetCommState(hParallelPort, MyDCB)
            ' If Success = False Then
            'Throw New CommException("Unable to retrieve the current control settings")
            'End If
            ' Modify the properties of MyDCB as appropriate.
            ' WARNING: Make sure to modify the properties in accordance with their supported values.
            'MyDCB.BaudRate = 9600
            'MyDCB.ByteSize = 8
            'MyDCB.Parity = NOPARITY
            'MyDCB.StopBits = ONESTOPBIT
            '' Reconfigure LPT1 based on the properties of MyDCB.
            'Success = SetCommState(hParallelPort, MyDCB)
            'If Success = False Then
            'Throw New CommException("Unable to reconfigure LPT1")
            'End If
            ' Retrieve the current time-out settings.
            'Success = GetCommTimeouts(hParallelPort, MyCommTimeouts)
            'If Success = False Then
            'Throw New CommException("Unable to retrieve current time-out settings")
            'End If
            ' Modify the properties of MyCommTimeouts as appropriate.
            ' WARNING: Make sure to modify the properties in accordance with their supported values.
            MyCommTimeouts.ReadIntervalTimeout = 0
            MyCommTimeouts.ReadTotalTimeoutConstant = 0
            MyCommTimeouts.ReadTotalTimeoutMultiplier = 0
            MyCommTimeouts.WriteTotalTimeoutConstant = 0
            MyCommTimeouts.WriteTotalTimeoutMultiplier = 0
            ' Reconfigure the time-out settings, based on the properties of MyCommTimeouts.
            'Success = SetCommTimeouts(hParallelPort, MyCommTimeouts)
            'If Success = False Then
            'Throw New CommException("Unable to reconfigure the time-out settings")
            'End If
            ' Write data to LPT1.
            '  Console.WriteLine(Chr(1))
            Success = WriteFile(hParallelPort, Buffer, Buffer.Length, BytesWritten, IntPtr.Zero)
            If Success = False Then
                '  Throw New CommException("Imposible escribir en puerto paralelo")
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            ' Release the handle to LPT1.
            Success = CloseHandle(hParallelPort)
            If Success = False Then
                'MsgBox("Y luego por aki")
                '  Console.WriteLine("Unable to release handle to LPT1")
            End If

        End Try

        '   Console.WriteLine("Press ENTER to quit")
        'Console.ReadLine()
    End Sub


    Sub escribe_puerto2(ByVal mi_byte As Integer, ByVal puerto As String)
        Dim hSerialPort, hParallelPort As IntPtr
        Dim Success As Boolean
        Dim MyDCB As DCB
        Dim MyCommTimeouts As COMMTIMEOUTS
        Dim BytesWritten, BytesRead As Int32
        Dim Buffer() As Byte

        ' Declare variables to use for encoding.
        Dim oEncoder As New System.Text.ASCIIEncoding
        Dim oEnc As System.Text.Encoding = oEncoder.GetEncoding(1252)

        ' Convert String to Byte().
        Buffer = oEnc.GetBytes(Chr(mi_byte))

        Try
            ' Parallel port.
            'Console.WriteLine("Accessing the LPT1 parallel port")
            'Obtain a handle to the LPT1 parallel port.
            hParallelPort = CreateFile(puerto, GENERIC_READ Or GENERIC_WRITE, 0, IntPtr.Zero,
               OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)
            'Verify that the obtained handle is valid.
            If hParallelPort.ToInt32 = -1 Then
                'Throw New CommException("Unable to obtain a handle to the LPT1 port")
            End If
            ' Retrieve the current control settings.
            '  Success = GetCommState(hParallelPort, MyDCB)
            ' If Success = False Then
            'Throw New CommException("Unable to retrieve the current control settings")
            'End If
            ' Modify the properties of MyDCB as appropriate.
            ' WARNING: Make sure to modify the properties in accordance with their supported values.
            'MyDCB.BaudRate = 9600
            'MyDCB.ByteSize = 8
            'MyDCB.Parity = NOPARITY
            'MyDCB.StopBits = ONESTOPBIT
            '' Reconfigure LPT1 based on the properties of MyDCB.
            'Success = SetCommState(hParallelPort, MyDCB)
            'If Success = False Then
            'Throw New CommException("Unable to reconfigure LPT1")
            'End If
            ' Retrieve the current time-out settings.
            'Success = GetCommTimeouts(hParallelPort, MyCommTimeouts)
            'If Success = False Then
            'Throw New CommException("Unable to retrieve current time-out settings")
            'End If
            ' Modify the properties of MyCommTimeouts as appropriate.
            ' WARNING: Make sure to modify the properties in accordance with their supported values.
            'MyCommTimeouts.ReadIntervalTimeout = 0
            'MyCommTimeouts.ReadTotalTimeoutConstant = 0
            'MyCommTimeouts.ReadTotalTimeoutMultiplier = 0
            'MyCommTimeouts.WriteTotalTimeoutConstant = 0
            'MyCommTimeouts.WriteTotalTimeoutMultiplier = 0
            ' Reconfigure the time-out settings, based on the properties of MyCommTimeouts.
            'Success = SetCommTimeouts(hParallelPort, MyCommTimeouts)
            'If Success = False Then
            'Throw New CommException("Unable to reconfigure the time-out settings")
            'End If
            ' Write data to LPT1.
            '  Console.WriteLine(Chr(1))
            Success = WriteFile(hParallelPort, Buffer, Buffer.Length, BytesWritten, IntPtr.Zero)
            ''   MsgBox(Success.ToString)
            If Success = False Then
                MsgBox("AKi fallo")
                'Throw New CommException("Imposible escribir en " & puerto)
            End If
            ' MsgBox("Llego hasta el final")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            MsgBox("fallo")
        Finally
            ' Release the handle to LPT1.
            Success = CloseHandle(hParallelPort)
            If Success = False Then
                MsgBox("Y luego por aki")
                '  Console.WriteLine("Unable to release handle to LPT1")
            End If

        End Try

        '   Console.WriteLine("Press ENTER to quit")
        'Console.ReadLine()

    End Sub
End Module
