'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

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

   'This procedure calculates the specified MS-DOS executable's header size including relocation items.
   ''--->>> THIS PROCEDURE HAS NOT YET BEEN IMPLEMENTED!
   Private Function GetEXEHeaderSize(Data As List(Of Byte)) As Integer
      Try
         Dim RelocationCount As Integer = BitConverter.ToInt16(Data.ToArray(), MSDOSHeaderE.RelocationCount)
         Dim RelocationTable As Integer = BitConverter.ToInt16(Data.ToArray(), MSDOSHeaderE.RelocationTableOffset)

         Return RelocationTable + (RelocationCount * &H4%)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

End Module
