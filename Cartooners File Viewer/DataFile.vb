'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Generic

'This class defines a data file.
Partial Public Class DataFileClass
   'This structure defines a data file.
   Public Structure DataFileStr
      Public Data As List(Of Byte)  'The file's data.
      Public Path As String         'The file's path.
   End Structure

   'This procedure defines a data file's export procedure.
   Public Sub Export(ExportPath As String)
   End Sub
End Class
