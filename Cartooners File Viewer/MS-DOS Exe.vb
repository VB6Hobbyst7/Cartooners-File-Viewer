'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Environment
Imports System.Text

'This module contains this program's core procedures.
Public Module MSDOSEXEModule
   'This enumeration contains the location of known data inside an MS-DOS executable's header.
   Public Enum MSDOSHeaderE As Integer
      Checksum = &H12%                 'The executable's negative pgm checksum.
      CodeSegment = &H16%              'The code segment (CS) register's initial value.
      HeaderSize = &H8%                'The executable's header size in paragraphs of 0x10 bytes.
      ImageSize = &H4%                 'The executable's image size in pages of 0x200 bytes.
      ImageSizeModulo = &H2%           'The executable's image size modulo (of 0x200).
      InstructionPointer = &H14%       'The instruction pointer (IP) register's initial value.
      MaximumParagraphs = &HC%         'The executable's maximum memory requirement in paragraphs of 0x10 bytes.
      MinimumParagraphs = &HA%         'The executable's minimum memory requirement in paragraphs of 0x10 bytes.
      OverlayNumber = &H20%            'The overlay number.
      RelocationCount = &H6%           'The executable's number or relocation items.
      RelocationTableOffset = &H18%    'The relocation table's offset.
      Signature = &H0%                 'The MS-DOS executable's signature.
      StackPointer = &H10%             'The stack pointer (SP) register's initial value.
      StackSegment = &HE%              'The stack segment (SS) register's initial value.
   End Enum

   'This structure defines an MS-DOS executable's header.
   Public Structure MSDOSHeaderStr
      Public Signature As Integer              'The MS-DOS executable's signature.
      Public ImageSizeModulo As Integer        'The executable's image size modulo (of 0x200).
      Public ImageSize As Integer              'The executable's image size in pages of 0x200 bytes.
      Public RelocationCount As Integer        'The executable's number or relocation items.
      Public HeaderSize As Integer             'The executable's header size in paragraphs of 0x10 bytes.
      Public MinimumParagraphs As Integer      'The executable's maximum memory requirement in paragraphs of 0x10 bytes.
      Public MaximumParagraphs As Integer      'The executable's minimum memory requirement in paragraphs of 0x10 bytes.
      Public StackSegment As Integer           'The stack segment (SS) register's initial value.
      Public StackPointer As Integer           'The stack pointer (SP) register's initial value.
      Public Checksum As Integer               'The executable's negative pgm checksum.
      Public InstructionPointer As Integer     'The instruction pointer (IP) register's initial value.
      Public CodeSegment As Integer            'The code segment (CS) register's initial value.
      Public RelocationTableOffset As Integer  'The relocation table's offset.
      Public OverlayNumber As Integer          'The overlay number.
   End Structure

   Public Const MSDOS_EXECUTABLE_SIGNATURE As Integer = &H5A4D%   'Contains the MS-DOS executable signature "MZ".
   Public Const MSDOS_HEADER_SIZE As Integer = &H1C%              'Contains the MS-DOS exectuable header's size.

   'This procedure manages a MS-DOS executable's header size with the size of the relocation table added.
   Public Function EXEHeaderSize(Optional NewData As List(Of Byte) = Nothing) As Integer
      Try
         Static CurrentExeHeaderSize As New Integer

         If NewData IsNot Nothing Then
            CurrentExeHeaderSize = BitConverter.ToUInt16(NewData.ToArray(), MSDOSHeaderE.RelocationTableOffset) + (BitConverter.ToUInt16(NewData.ToArray(), MSDOSHeaderE.RelocationCount) * &H4%)
         End If

         Return CurrentExeHeaderSize
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified executable's header and relocation table information as text.
   Public Function GetEXEHeaderInformation(Data As List(Of Byte)) As String
      Dim NewText As New StringBuilder
      Dim RelocationCount As Integer = BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.RelocationCount)
      Dim RelocationTableOffset As Integer = BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.RelocationTableOffset)

      With Data
         Try
            NewText.Append($"Signature: ""{GetString(Data, MSDOSHeaderE.Signature, &H2%)}""{NewLine}")
            NewText.Append($"Image size modulo (of 0x200): 0x{ BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.ImageSizeModulo):X} bytes.{NewLine}")
            NewText.Append($"Image size in pages (of 0x200 bytes): 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.ImageSize):X}.{NewLine}")
            NewText.Append($"Number of relocation items: 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.RelocationCount):X}.{NewLine}")
            NewText.Append($"Header sizes in paragraphs (of 0x10 bytes): 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.HeaderSize):X}.{NewLine}")
            NewText.Append($"Minimum number of paragraphs required (of 0x10 bytes): 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.MinimumParagraphs):X}.{ NewLine}")
            NewText.Append($"Maximum number of paragraphs required (of 0x10 bytes): 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.MaximumParagraphs):X}.{NewLine}")
            NewText.Append($"Stack segment (SS) register: 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.StackSegment):X}.{NewLine}")
            NewText.Append($"Stack pointer (SP) register: 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.StackPointer):X}.{NewLine}")
            NewText.Append($"Negative checksum of PGM: 0x{ BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.Checksum):X}.{NewLine}")
            NewText.Append($"Instruction pointer (IP) register: 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.InstructionPointer):X}.{NewLine}")
            NewText.Append($"Code segment (CS) register: 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.CodeSegment):X}.{ NewLine}")
            NewText.Append($"Relocation table offset: 0x{RelocationTableOffset:X}.{NewLine}")
            NewText.Append($"Overlay number: 0x{BitConverter.ToUInt16(Data.ToArray(), MSDOSHeaderE.OverlayNumber):X}.{NewLine}")
            NewText.Append(NewLine)
            NewText.Append($"Relocation items:{NewLine}")

            RelocationItems().ForEach(Sub(RelocationItem As Integer) NewText.Append($"{RelocationItem:X8}{NewLine}"))

            Return NewText.ToString()
         Catch ExceptionO As Exception
            HandleError(ExceptionO)
         End Try

         Return Nothing
      End With
   End Function

   'This procedure manages the current executable's relocation items.
   Public Function RelocationItems(Optional NewData As List(Of Byte) = Nothing) As List(Of Integer)
      Try
         Dim RelocationCount As New Integer
         Dim RelocationTableOffset As New Integer
         Static CurrentRelocationItems As New List(Of Integer)

         If NewData IsNot Nothing Then
            CurrentRelocationItems.Clear()
            RelocationCount = BitConverter.ToUInt16(NewData.ToArray(), MSDOSHeaderE.RelocationCount)
            RelocationTableOffset = BitConverter.ToUInt16(NewData.ToArray(), MSDOSHeaderE.RelocationTableOffset)

            For Position As Integer = RelocationTableOffset To RelocationTableOffset + ((RelocationCount - &H1%) * &H4%) Step &H4%
               CurrentRelocationItems.Add(BitConverter.ToInt32(NewData.ToArray, Position))
            Next Position
         End If

         Return CurrentRelocationItems
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
