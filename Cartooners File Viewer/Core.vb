'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Convert
Imports System.Drawing
Imports System.Environment
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Math
Imports System.Text
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module CoreModule
   'This enumeration lists the supported file types.
   Public Enum FileTypesE As Integer
      None                    'No file.
      Actor                   'Actor files.
      Archive                 'Archive files.
      Executable              'Executable files.
      LBM                     'LBM files.
      Movie                   'Movie files.
      Music                   'Music files.
      Preferences             'Preferences files.
      Script                  'Script files.
   End Enum

   'This enumeration contains the location of known data inside an MS-DOS executable's header.
   Public Enum MSDOSHeaderE As Integer
      Checksum = &H12%                  'The executable's negative pgm checksum.
      CodeSegment = &H16%               'The code segment (CS) register's initial value.
      HeaderSize = &H8%                 'The executable's header size in paragraphs of 0x10 bytes.
      ImageSize = &H4%                  'The executable's image size in pages of 0x200 bytes.
      ImageSizeModulo = &H2%            'The executable's image size modulo (of 0x200).
      InstructionPointer = &H14%        'The instruction pointer (IP) register's initial value.
      MaximumParagraphs = &HC%          'The executable's maximum memory requirement in paragraphs of 0x10 bytes.
      MinimumParagraphs = &HA%          'The executable's minimum memory requirement in paragraphs of 0x10 bytes.
      OverlayNumber = &H20%             'The overlay number.
      RelocationCount = &H6%            'The executable's number or relocation items.
      RelocationTableOffset = &H18%     'The relocation table's offset.
      Signature = &H0%                  'The MS-DOS executable's signature.
      StackPointer = &H10%              'The stack pointer (SP) register's initial value.
      StackSegment = &HE%               'The stack segment (SS) register's initial value.
   End Enum

   'This enumaration lists the nibbles in a byte.
   Public Enum NibblesE As Integer
      HighNibble             'The high nibble.
      LowNibble              'The low nibble.
   End Enum

   'This structure defines an LZW dictionary entry.
   Public Structure LZWEntryStr
      Public Prefix As Integer   'Contains an LZW entry's prefix.
      Public Suffix As Integer   'Contains an LZW entry's suffix.
   End Structure

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

   'The core constants used by this program.
   Public Const ACTOR_TEMPLATE As String = "Actor"                'Contains the identifier for actor templates.
   Public Const DELIMITER As Char = ControlChars.NullChar         'Contains the delimiter used in various types of data.
   Public Const GBR_12_COLOR_DEPTH As Integer = &H10%             'Contains the number of colors in a 12bit GBR palette.
   Public Const GBR_12_COLOR_LENGTH As Integer = &H2%             'Contains the number of bytes per color in a 12bit GBR palette.
   Public Const LZW_END As Integer = &H101%                       'Contains the end of a LZW value sequence.
   Public Const LZW_MAXIMUM_BITS As Integer = &HC%                'Contains the maximum number of bits per value in a LZW sequence.
   Public Const LZW_START As Integer = &H100%                     'Contains the start of a LZW value sequence.
   Public Const LZW_SYMBOL_BASE As Integer = &H102%               'Contains the lowest value used for an LZW symbol.
   Public Const LZW_SYMBOL_TOP As Integer = &HFFF%                'Contains the highest value used for an LZW symbol.
   Public Const MSDOS_EXECUTABLE_SIGNATURE As Integer = &H5A4D%   'Contains the MS-DOS executable signature "MZ".
   Public Const MSDOS_HEADER_SIZE As Integer = &H1C%              'Contains the MS-DOS exectuable header's size.
   Public Const NOT_FOUND As Integer = -1                         'Indicates that a value could not be found in a given data set.
   Public Const PADDING As Char = ControlChars.NullChar           'Contains the character used to pad a path.
   Public Const PAGE_SIZE As Integer = &H200%                     'Contains the size of a page in memory.
   Public Const PARAGRAPH_SIZE As Integer = &H10%                 'Contains the size of a paragraph in memory.
   Public Const PIXELS_PER_BYTE As Integer = &H2%                 'Contains the number pixels per byte in an uncompressed image.
   Public Const SCRIPT_TEMPLATE As String = "Script"              'Contains the identifier for actor templates.

   Public ReadOnly ARGB_TO_GBR As Func(Of Color, Byte()) = Function(ARGB As Color) {ToByte(ToInt32(ARGB.G >> &H4%) << &H4% Or ToInt32(ARGB.B >> &H4%)), ToByte(ARGB.R >> &H4%)}                                                                      'This procedure converts the specified 24 bit ARGB color to a 12 bit GBR color.
   Public ReadOnly BYTES_TO_TEXT As Func(Of List(Of Byte), String) = Function(Bytes As List(Of Byte)) New String((From ByteO In Bytes Select ToChar(ByteO)).ToArray())                                                                               'This procedure converts the specified bytes to text.
   Public ReadOnly COLOR_DIFFERENCE As Func(Of Color, Color, Integer) = Function(Color1 As Color, Color2 As Color) CInt((Abs(CInt(Color2.R) - CInt(Color1.R)) + Abs(CInt(Color2.G) - CInt(Color1.G)) + Abs(CInt(Color2.B) - CInt(Color1.B))) / 3)    'This procedure returns the difference between the two specified colors.
   Public ReadOnly GET_BIT As Func(Of Byte, Integer, Integer) = Function(ByteO As Byte, BitIndex As Integer) Abs(ToInt32((New BitArray({ByteO}))(BitIndex)))                                                                                         'This procedure returns the specified bit inside the specified byte.
   Public ReadOnly GET_DWORD As Func(Of List(Of Byte), Integer, Integer) = Function(Data As List(Of Byte), Position As Integer) (BitConverter.ToInt32(Data.ToArray(), Position))                                                                     'This procedure extracts a little endian DWORD value from the specified bytes at the specified position and returns it.
   Public ReadOnly GET_WORD As Func(Of List(Of Byte), Integer, Integer) = Function(Data As List(Of Byte), Position As Integer) (BitConverter.ToInt16(Data.ToArray(), Position))                                                                      'This procedure extracts a little endian WORD value from the specified bytes at the specified position and returns it.
   Public ReadOnly LZW_MAXIMUM_ENTRIES As Integer = (&H1% << LZW_MAXIMUM_BITS)                                                                                                                                                                       'Contains the maximum number of LZW symbols possible with the maximum LZW bit count.
   Public ReadOnly SET_BIT As Func(Of Integer, Integer, Integer, Byte) = Function(ByteO As Integer, BitIndex As Integer, Bit As Integer) CByte(ByteO Or (Bit << BitIndex))                                                                           'This procedure sets the specified bit inside the specified byte.
   Public ReadOnly TERMINATE_AT_NULL As Func(Of String, String) = Function(Text As String) If(Text.Contains(ControlChars.NullChar), Text.Substring(0, Text.IndexOf(ControlChars.NullChar)), Text)                                                    'This procedure terminates the specified text at the left most null character and returns the result.
   Public ReadOnly TEXT_TO_BYTES As Func(Of String, List(Of Byte)) = Function(Text As String) (From Character In Text.ToCharArray() Select ToByte(Character)).ToList()                                                                               'This procedure converts the specified text to bytes.
   Public ReadOnly TOGGLE_WORD As Func(Of List(Of Byte), Integer) = Function(Bytes As List(Of Byte)) ToInt32(String.Format("{0:X2}{1:X2}", Bytes(&H1%), Bytes(&H0%)), fromBase:=16)                                                                  'This procedure reverses the specified word's byte order.
   Public ReadOnly UNSIGN_BYTE As Func(Of Integer, Integer) = Function(Value As Integer) Abs(If(Value >= &H80%, Value - &H100%, Value))                                                                                                              'This procedure converts a signed byte to an unsigned byte.

   'This function converts MS-DOS line breaks to Microsoft Windows line breaks in the specified text and returns the result.
   Public Function ConvertMSDOSLineBreak(Text As String) As String
      Try
         Dim Conversion As New StringBuilder
         Dim Position As Integer = 0

         Do While Position < Text.Length
            If Position < Text.Length - NewLine.Length AndAlso Text.Substring(Position, NewLine.Length) = NewLine Then
               Conversion.Append(NewLine)
               Position += 1
            ElseIf Text.Chars(Position) = ControlChars.Cr Then
               Conversion.Append(NewLine)
            Else
               Conversion.Append(Text.Chars(Position))
            End If
            Position += 1
         Loop

         Return Conversion.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure draws the specified 4 bit pixel data on an image.
   Public Function Draw4BitImage(Data As List(Of Byte), Width As Integer, Height As Integer, Palette As List(Of Color), BytesPerRow As Integer, Optional TransparentIndex As Integer? = Nothing, Optional TransparentColor As Color = Nothing) As Bitmap
      Try
         Dim Color1 As New Integer
         Dim Color2 As New Integer
         Dim ImageO As New Bitmap(Width, Height)
         Dim Position As Integer = &H0%

         For y As Integer = &H0% To Height - &H1%
            For x As Integer = &H0% To (BytesPerRow * PIXELS_PER_BYTE) - &H1% Step PIXELS_PER_BYTE
               Color1 = GetNibble(Data(Position), NibblesE.HighNibble)
               Color2 = GetNibble(Data(Position), NibblesE.LowNibble)
               If TransparentIndex Is Nothing Then
                  ImageO.SetPixel(x, y, Palette(Color1))
                  If x + &H1% < Width Then ImageO.SetPixel(x + &H1%, y, Palette(Color2))
               Else
                  ImageO.SetPixel(x, y, If(Color1 = TransparentIndex, TransparentColor, Palette(Color1)))
                  If x + &H1% < Width Then ImageO.SetPixel(x + &H1%, y, If(Color2 = TransparentIndex, TransparentColor, Palette(Color2)))
               End If

               Position += &H1%
            Next x
         Next y

         Return ImageO
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts non-displayable or all (if specified) characters in the specified text to escape sequences.
   Public Function Escape(ToEscape As Object, Optional EscapeCharacter As Char = "/"c, Optional EscapeAll As Boolean = False) As String
      Try
         Dim Character As New Char
         Dim Escaped As New StringBuilder
         Dim Index As Integer = 0
         Dim Text As String = If(TypeOf ToEscape Is List(Of Byte), BYTES_TO_TEXT(DirectCast(ToEscape, List(Of Byte))), DirectCast(ToEscape, String))

         With Escaped
            Do Until Index >= Text.Length
               Character = Text.Chars(Index)

               If Character = EscapeCharacter AndAlso Not EscapeAll Then
                  .Append(New String(EscapeCharacter, 2))
               ElseIf (Character = ControlChars.Tab OrElse Character >= " "c) AndAlso Not EscapeAll Then
                  .Append(Character)
               ElseIf (Index < Text.Length - 1 AndAlso Character & Text.Chars(Index + 1) = NewLine) AndAlso Not EscapeAll Then
                  .Append(NewLine)
                  Index += 1
               Else
                  .Append($"{EscapeCharacter}{ToByte(Character):X2}")
               End If
               Index += 1
            Loop
         End With

         Return Escaped.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure searches the specified bytea array for the specified bytes and returns their position if found.
   Public Function FindBytes(Bytes As List(Of Byte), SearchBytes As List(Of Byte), Offset As Integer) As Integer
      Try
         For Position As Integer = Offset To (Bytes.Count - 1) - SearchBytes.Count
            If GetBytes(Bytes, Position, SearchBytes.Count).SequenceEqual(SearchBytes) Then Return Position
         Next Position

         Return NOT_FOUND
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure reads the GBR palette from the specified data at the specified location.
   Public Function GBRPalette(Data As List(Of Byte), PaletteLocation As Integer) As List(Of Color)
      Try
         Dim Palette As New List(Of Color)

         For Position As Integer = PaletteLocation To PaletteLocation + ((GBR_12_COLOR_DEPTH - &H1%) * GBR_12_COLOR_LENGTH) Step GBR_12_COLOR_LENGTH
            Palette.Add(GBRToARGB(New List(Of Byte)({Data(Position), Data(Position + &H1%)})))
         Next Position

         Return Palette
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified 12 bit GBR color to a 24 bit ARGB color.
   Public Function GBRToARGB(GBR As List(Of Byte)) As Color
      Try
         Dim Blue As Integer = GetNibble(GBR(&H0%), NibblesE.LowNibble)
         Dim Green As Integer = GetNibble(GBR(&H0%), NibblesE.HighNibble)
         Dim Red As Integer = GetNibble(GBR(&H1%), NibblesE.LowNibble)

         If Not GetNibble(GBR(&H1%), NibblesE.HighNibble) = &H0% Then
            MessageBox.Show("Invalid GBR color value.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
         End If

         Return Color.FromArgb((Red << &H4%) Or Red, (Green << &H4%) Or Green, (Blue << &H4%) Or Blue)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified palettes' data to hexadecimal output with descriptions.
   Public Function GBRToText(Header As String, Palettes As List(Of List(Of Color)), Optional Descriptions As List(Of String) = Nothing) As String
      Try
         Dim GBRText As New StringBuilder

         GBRText.Append($"{Header}{NewLine}{NewLine}")
         For Palette As Integer = 0 To Palettes.Count - 1
            If Descriptions IsNot Nothing Then GBRText.Append($"Palette - {Descriptions(Palette)}:{NewLine}")
            GBRText.Append($"I: R: G: B:{NewLine}")
            For Index As Integer = 0 To Palettes(Palette).Count - 1
               With Palettes(Palette)(Index)
                  GBRText.Append($"{Index,2}")
                  GBRText.Append($"{$"{ .R:X2}",3}")
                  GBRText.Append($"{$"{ .G:X2}",3}")
                  GBRText.Append($"{$"{ .B:X2}",3}{NewLine}")
               End With
            Next Index
            GBRText.Append(NewLine)
         Next Palette

         Return GBRText.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified number of bytes at the specified position.
   Public Function GetBytes(Data As List(Of Byte), ByRef Offset As Integer, Count As Integer, Optional AdvanceOffset As Boolean = False) As List(Of Byte)
      Try
         Dim Bytes As New List(Of Byte)(Data.GetRange(Offset, Count))

         If AdvanceOffset Then Offset += Count

         Return Bytes
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure loads the specified template and returns the appropriate data file type.
   Public Function GetDataFileFromTemplate(ByRef PathO As String, DataFileMenu As ToolStripMenuItem) As DataFileClass
      Try
         For Each Line As String In LoadTemplate(PathO)
            Select Case Line.Trim().ToLower()
               Case $"[{ACTOR_TEMPLATE.ToLower()}]"
                  Return New ActorClass(PathO, DataFileMenu)
               Case $"[{SCRIPT_TEMPLATE.ToLower()}]"
                  Return New ScriptClass(PathO, DataFileMenu)
               Case Else
                  If Not Line.Trim().ToLower() = Nothing Then
                     MessageBox.Show("This template type is not recognized.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                     Return Nothing
                  End If
            End Select
         Next Line
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified nibble at the specified position.
   Public Function GetNibble(ByteO As Byte, Nibble As NibblesE) As Integer
      Try
         Select Case Nibble
            Case NibblesE.HighNibble
               Return (ByteO And &HF0%) >> &H4%
            Case NibblesE.LowNibble
               Return ByteO And &HF%
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns a string of the length specified in bytes at the specified position.
   Public Function GetString(Data As List(Of Byte), ByRef Offset As Integer, Count As Integer, Optional AdvanceOffset As Boolean = False) As String
      Try
         Return BYTES_TO_TEXT(GetBytes(Data, Offset, Count, AdvanceOffset))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procudure handles any errors that occur.
   Public Sub HandleError(ExceptionO As Exception)
      Try
         If MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Error) = DialogResult.Cancel Then Application.Exit()
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure loads the specified template and returns it without comment lines.
   Public Function LoadTemplate(Optional PathO As String = Nothing) As List(Of String)
      Try
         Static Template As List(Of String) = Nothing

         If Not PathO = Nothing Then Template = New List(Of String)(From Item As String In File.ReadAllLines(PathO) Where Not Item.Trim().StartsWith("#"))

         Return Template
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified bytes to a number.
   Public Function NumberFromBytes(Bytes As List(Of Byte)) As Integer
      Try
         Dim Hexadecimals As New StringBuilder

         Bytes.ForEach(Sub(ByteO As Byte) Hexadecimals.Append($"{ByteO:X2}"))

         Return ToInt32(Hexadecimals.ToString(), fromBase:=16)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified number to a series of bytes of the specified length.
   Public Function NumberToBytes(Number As Integer, Length As Integer) As List(Of Byte)
      Try
         Select Case Length
            Case &H2%
               Return BitConverter.GetBytes(ToUInt16(Number)).Reverse.ToList()
            Case &H4%
               Return BitConverter.GetBytes(ToUInt32(Number)).Reverse.ToList()
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns information about this program.
   Public Function ProgramInformation() As String
      Try
         With My.Application.Info
            Return $"{ .Title} v{ .Version} - by: { .CompanyName}"
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure sets the specified nibble at the specified position.
   Public Function SetNibble(ByteO As Integer, NewNibble As Integer, Nibble As NibblesE) As Integer
      Try
         Select Case Nibble
            Case NibblesE.HighNibble
               Return ByteO Or (NewNibble << &H4%)
            Case NibblesE.LowNibble
               Return ByteO Or NewNibble
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts an unsigned byte to a signed byte.
   Public Function SignByte(Value As Integer, IsNegative As Boolean) As Integer
      Try
         Return If(IsNegative, &H100% - Value, Value)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified number to a byte array containing a big endian number to a little endian number.
   Public Function ToLittleEndian(Number As Integer, Length As Integer) As List(Of Byte)
      Try
         Dim Hexadecimals As String = $"{Number:X}".PadLeft(Length * &H2%, "0"c)
         Dim LittleEndian As New List(Of Byte)

         For ByteO As Integer = Hexadecimals.Length - &H2% To &H0% Step -&H2%
            LittleEndian.Add(ToByte(Hexadecimals.Substring(ByteO, &H2%), fromBase:=16))
         Next ByteO

         Return LittleEndian
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure updates the interface window's databox. 
   Public Sub UpdateDataBox(Optional NewText As String = Nothing, Optional NewDataBox As TextBox = Nothing)
      Try
         Static CurrentDataBox As TextBox = Nothing

         If NewDataBox IsNot Nothing Then
            CurrentDataBox = NewDataBox
         ElseIf Not NewText = Nothing AndAlso CurrentDataBox IsNot Nothing Then
            With CurrentDataBox
               .Text = NewText
               .Select(0, 0)
               If .TextLength < NewText.Length Then MessageBox.Show("The databox cannot display all new data.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End With
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure converts any escape sequences in the specified text to characters.
   Public Function Unescape(Text As String, Optional EscapeCharacter As Char = "/"c, Optional ByRef ErrorAt As Integer = 0) As String
      Try
         Dim Character As New Char
         Dim CharacterCode As New Integer
         Dim Index As Integer = 0
         Dim NextCharacter As New Char
         Dim Unescaped As New StringBuilder

         ErrorAt = 0

         Do Until Index >= Text.Length OrElse ErrorAt > 0
            Character = Text.Chars(Index)
            If Index < Text.Length - 1 Then NextCharacter = Text.Chars(Index + 1) Else NextCharacter = Nothing

            If Character = EscapeCharacter Then
               If NextCharacter = EscapeCharacter Then
                  Unescaped.Append(Character)
                  Index += 1
               Else
                  If NextCharacter = Nothing Then
                     ErrorAt = Index + 1
                  Else
                     If Index < Text.Length - 2 AndAlso Integer.TryParse(Text.Substring(Index + 1, 2), NumberStyles.HexNumber, Nothing, CharacterCode) Then
                        Unescaped.Append(ToChar(CharacterCode))
                        Index += 2
                     Else
                        ErrorAt = Index + 1
                     End If
                  End If
               End If
            Else
               Unescaped.Append(Character)
            End If
            Index += 1
         Loop

         Return Unescaped.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
