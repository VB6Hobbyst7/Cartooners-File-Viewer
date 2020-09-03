'This class' imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Environment
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

'This class contains the movie related procedures.
Public Class ExecutableClass
   Inherits DataFileClass

   'This enumeration contains the location of all known data used by Cartooners.
   Private Enum DataLocationsE As Integer
      About = &H35AC4%                        'About text.
      ANewActor = &H3777F%                    'A new actor.
      ButtonEdgeStyle = &H35146%              'Button edge style information.
      C_FILE_INFO = &H39437%                  'C_FILE_INFO.
      ClearErase = &H3506A%                   'Clear/erase text.
      DaysMonths = &H39DE4%                   'Day and month abbreviations.
      DefaultSpeechBalloon = &H29045%         'Default speech balloon information.
      Dialog = &H36874%                       'Dialog text.
      DriveLetters = &H3516C%                 'Drive letters.
      ExclamationBalloon = &H29A39%           'Exclamation speech balloon information.
      ExecutableCode = &H2DD0%                '80286+ executable code.
      ExecutableExtensions = &H39D2C%         'Executable extensions.
      FileCharacters = &H3559F%               'File system related characters.
      FileExtensions = &H34BD3%               'File extensions.
      InvalidActor = &H376EA%                 'Invalid actor message.
      KeyboardLayout = &H37674%               'Keyboard layout.
      LoadSave = &H34B76%                     'Load/save text.
      Messages1 = &H34C43%                    'First messages.
      Messages2 = &H35E1C%                    'Second messages.
      Messages3 = &H35658%                    'Third messages.      
      Menus = &H34D5E%                        'Menus.
      MicrosoftRuntimeLibrary = &H33DE8%      'Microsoft Run-time Library text.
      Printing1 = &H36FA8%                    'First printing text.
      Printing2 = &H36FB9%                    'Second printing text.
      RunTimeErrors = &H39E46%                'Run-time error messages.
      SpeechBalloonLayouts = &H34A2C%         'Speech balloon layouts.
      SpeechBalloonSlotsWayMenu = &H2BED6%    'Speech balloon slots/way menu items.
      SpeechBalloonTypes = &H34964%           'Speech balloon types.
      Text1 = &H34A62%                        'First general text fragment.
      Text2 = &H350A4%                        'Second general text fragment.
      Text3 = &H35622%                        'Third general text fragment.
      Text4 = &H3598C%                        'Fourth general text fragment.
      Text5 = &H359E2%                        'Fifth general text fragment.
      Text6 = &H3774A%                        'Sixth general text fragment.
      TimeZones = &H39DCA%                    'The time zones.
      ThoughtSpeechBalloon = &H2947C%         'Thought balloon speech information.
      UnknownSpeechBalloonData = &H29040%     'Unknown speech balloon data.
      WayMenu = &H379C9%                      'The way menu items.
   End Enum

   'This enumeration contains the locations of all icons used by Cartooners.
   Private Enum IconLocationsE As Integer
      CheckedCircle = &H340BC%        'The checked circle icon.
      Circle = &H34074%               'The circle icon.
      CrossArrows = &H34520%          'The cross arrows icon.
      Eraser = &H34370%               'The eraser icon.
      FilledCircle = &H34104%         'The filled circle icon.
      LeftArrow = &H34208%            'The left arrow icon.
      LoopArrow = &H34490%            'The loop arrow icon.
      RightArrow = &H342BC%           'The right arrow icon.
      SpeechBalloon = &H34400%        'The speech balloon icon.
   End Enum

   'This enumeration contains the location of all images used by Cartooners.
   Private Enum ImageLocationsE As Integer
      MenuScreen = &H2C830%            'The run-length compressed menu screen image.
      TicketScreen = &H32330%          'The run-length compressed ticket screen image.
      TitleScreen = &H2F952%           'The run-length compressed title screen image.
   End Enum

   'This enumeration contains the location of all palettes used by Cartooners.
   Private Enum PaletteLocationsE As Integer
      GUICGAEGA = &H34C20%             'The GUI's CGA/EGA palette.
      MenuScreenCGA = &H2C7F0%         'The menu screen's CGA palette.
      MenuScreenEGA = &H2C810%         'The menu screen's EGA palette.
      TicketScreenCGA = &H32310%       'The ticket screen CGA palette.
      TicketScreenEGA = &H322F0%       'The ticket screen EGA palette.
      TitleScreenCGA = &H2F910%        'The title screen's CGA palette.
      TitleScreenEGA = &H2F930%        'The title screen's EGA palette.
   End Enum

   'This enumeration contains a list of the palettes used by Cartooners.
   Private Enum PalettesE As Integer
      GUICGAEGA = &H6%         'The GUI's CGA/EGA palette.
      MenuScreenCGA = &H0%     'The menu screen's CGA palette.
      MenuScreenEGA = &H1%     'The menu screen's EGA palette.
      TicketScreenCGA = &H5%   'The ticket screen's CGA palette.
      TicketScreenEGA = &H4%   'The ticket screen's EGA palette.
      TitleScreenCGA = &H2%    'The title screen's CGA palette.
      TitleScreenEGA = &H3%    'The title screen's EGA palette.
   End Enum

   Private Const BYTES_PER_ROW As Integer = &HA0%                  'Contains the number of bytes per pixel row.
   Private Const EXPECTED_NAME As String = "Cartoons.exe"          'Contains the executable's expected file name.
   Private Const EXPECTED_PACKED_SIZE As Integer = &H36A5F%        'Contains the executable's expected packed file size.
   Private Const EXPECTED_UNPACKED_SIZE As Integer = &H39F20%      'Contains the executable's expected unpacked file size.
   Private Const SCREEN_HEIGHT As Integer = &HC8%                  'Contains the screen height used by Cartooners in pixels.
   Private Const SCREEN_WIDTH As Integer = &H140%                  'Contains the screen width used by Cartooners in pixels.

   Private ReadOnly DATA_SIZES As New Dictionary(Of Integer, Integer) From
      {{DataLocationsE.About, &HE2%},
       {DataLocationsE.ANewActor, &HC%},
       {DataLocationsE.ButtonEdgeStyle, &H26%},
       {DataLocationsE.C_FILE_INFO, &HB%},
       {DataLocationsE.ClearErase, &H2D%},
       {DataLocationsE.DaysMonths, &H3A%},
       {DataLocationsE.DefaultSpeechBalloon, &H437%},
       {DataLocationsE.Dialog, &H316%},
       {DataLocationsE.DriveLetters, &H54%},
       {DataLocationsE.ExclamationBalloon, &H801%},
       {DataLocationsE.ExecutableCode, &H26240%},
       {DataLocationsE.ExecutableExtensions, &HF%},
       {DataLocationsE.FileCharacters, &H17%},
       {DataLocationsE.FileExtensions, &H1D%},
       {DataLocationsE.InvalidActor, &H35%},
       {DataLocationsE.KeyboardLayout, &H75%},
       {DataLocationsE.LoadSave, &H3E%},
       {DataLocationsE.Messages1, &H11B%},
       {DataLocationsE.Messages2, &H2C0%},
       {DataLocationsE.Messages3, &H25F%},
       {DataLocationsE.Menus, &H2C2%},
       {DataLocationsE.MicrosoftRuntimeLibrary, &H16B%},
       {DataLocationsE.Printing1, &HA%},
       {DataLocationsE.Printing2, &H22%},
       {DataLocationsE.RunTimeErrors, &HD2%},
       {DataLocationsE.SpeechBalloonLayouts, &H36%},
       {DataLocationsE.SpeechBalloonSlotsWayMenu, &H90C%},
       {DataLocationsE.SpeechBalloonTypes, &HC6%},
       {DataLocationsE.Text1, &H7A%},
       {DataLocationsE.Text2, &H1E%},
       {DataLocationsE.Text3, &H36%},
       {DataLocationsE.Text4, &H38%},
       {DataLocationsE.Text5, &HDE%},
       {DataLocationsE.Text6, &H42%},
       {DataLocationsE.ThoughtSpeechBalloon, &H5BD%},
       {DataLocationsE.TimeZones, &HB%},
       {DataLocationsE.UnknownSpeechBalloonData, &H5%},
       {DataLocationsE.WayMenu, &H103%}}  'The data sizes.

   Private ReadOnly GET_TEXT As Func(Of DataLocationsE, String) = Function(Location As DataLocationsE) GetString(DataFile().Data, CInt(Location), DATA_SIZES(Location))   'This procedure returns the text at the specified location.

   Private ReadOnly IMAGE_SIZES As New Dictionary(Of Integer, Integer) From
      {{ImageLocationsE.MenuScreen, &H262C%},
      {ImageLocationsE.TicketScreen, &H1A9C%},
      {ImageLocationsE.TitleScreen, &H2995%}}   'The image sizes.

   'The menu items used by this class.
   Private WithEvents DisplayDataMenu As New ToolStripMenuItem With {.Text = "Display &Data"}
   Private WithEvents DisplayDataSubmenu As New ToolStripComboBox
   Private WithEvents DisplayInformationMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F1, .Text = "Display &Information"}
   Private WithEvents DisplayPalettesMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F2, .Text = "Display &Palettes"}

   'This procedure initializes this class.
   Public Sub New(PathO As String, Optional DataFileMenu As ToolStripMenuItem = Nothing)
      Try
         Dim TextSubMenuItems As New List(Of String)

         If DataFile(ExecutablePath:=PathO).Data IsNot Nothing AndAlso DataFileMenu IsNot Nothing Then
            For Each Location As Integer In CType([Enum].GetValues(GetType(DataLocationsE)), Integer())
               TextSubMenuItems.Add(DirectCast(Location, DataLocationsE).ToString())
            Next Location
            TextSubMenuItems.Sort()

            With DisplayDataSubmenu
               .Items.Clear()
               TextSubMenuItems.ForEach(Sub(SubMenuItem As String) .Items.Add(New ToolStripMenuItem With {.CheckOnClick = True, .Text = SubMenuItem}))
            End With

            DisplayDataMenu.DropDownItems.Add(DisplayDataSubmenu)

            With DataFileMenu
               .DropDownItems.Clear()
               .DropDownItems.AddRange({DisplayInformationMenu, DisplayPalettesMenu, DisplayDataMenu})
               .Text = "&Cartooners"
               .Visible = True
            End With
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedures manages the executable's data file.
   Private Function DataFile(Optional ExecutablePath As String = Nothing) As DataFileStr
      Try
         Static CurrentFile As New DataFileStr With {.Data = Nothing, .Path = Nothing}

         With CurrentFile
            If Not ExecutablePath = Nothing Then
               If Path.GetFileName(ExecutablePath).ToUpper() = EXPECTED_NAME.ToUpper() Then
                  If New FileInfo(ExecutablePath).Length = EXPECTED_PACKED_SIZE Then
                     .Data = UnpackExecutable(New List(Of Byte)(File.ReadAllBytes(ExecutablePath)))
                     If .Data.Count = EXPECTED_UNPACKED_SIZE Then
                        If GET_WORD(.Data, MSDOSHeaderE.Signature) = MSDOS_EXECUTABLE_SIGNATURE Then
                           .Path = ExecutablePath
                           Palettes(Refresh:=True)

                           DisplayInformationMenu.PerformClick()
                        Else
                           MessageBox.Show("Invalid MS-DOS executable signature.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                           .Data = Nothing
                        End If
                     Else
                        MessageBox.Show($"An error occurred during unpacking. Expected unpacked size: {EXPECTED_UNPACKED_SIZE} bytes. Actual unpacked size: { .Data.Count} bytes.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        .Data = Nothing
                     End If
                  Else
                     MessageBox.Show($"Wrong size. Expected packed size: {EXPECTED_PACKED_SIZE} bytes.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                  End If
               Else
                  MessageBox.Show($"Wrong executable. Expected name: ""{EXPECTED_NAME}"".", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
               End If
            End If
         End With

         Return CurrentFile
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure displays the current executable's speech balloon slots/way menu.
   Private Sub DisplayDataSubmenu_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DisplayDataSubmenu.SelectedIndexChanged
      Try
         Dim NewText As New StringBuilder
         Dim SelectedLocation As New DataLocationsE

         For Each Location As Integer In CType([Enum].GetValues(GetType(DataLocationsE)), Integer())
            SelectedLocation = DirectCast(Location, DataLocationsE)
            If SelectedLocation.ToString() = DisplayDataSubmenu.SelectedItem.ToString() Then Exit For
         Next Location

         Select Case SelectedLocation
            Case DataLocationsE.ButtonEdgeStyle, DataLocationsE.DefaultSpeechBalloon, DataLocationsE.ExclamationBalloon, DataLocationsE.ThoughtSpeechBalloon, DataLocationsE.UnknownSpeechBalloonData
               NewText.Append($"{Escape(DataFile().Data, " "c, EscapeAll:=True).Trim()}")
            Case DataLocationsE.ExecutableCode
               NewText.Append($"Executable code:{NewLine}")
               NewText.Append($"{Escape(DataFile().Data, " "c, EscapeAll:=True).Trim()}")
            Case Else
               NewText.Append(Escape(ConvertMSDOSLineBreak(GET_TEXT(SelectedLocation)).Replace(DELIMITER, NewLine)))
         End Select

         UpdateDataBox(NewText.ToString())
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the current executable's information.
   Private Sub DisplayInformationMenu_Click(sender As Object, e As EventArgs) Handles DisplayInformationMenu.Click
      Try

         Dim NewText As New StringBuilder
         Dim RelocationCount As Integer = GET_WORD(DataFile().Data, MSDOSHeaderE.RelocationCount)
         Dim RelocationTableOffset As Integer = GET_WORD(DataFile().Data, MSDOSHeaderE.RelocationTableOffset)

         With DataFile()
            NewText.Append($"{ .Path}:{NewLine}{NewLine}")
            NewText.Append($"File size: {New FileInfo(.Path).Length} bytes.{ NewLine}")
            NewText.Append($"Signature: ""{GetString(.Data, MSDOSHeaderE.Signature, &H2%)}""{NewLine}")
            NewText.Append($"Image size modulo (of 0x200): 0x{ GET_WORD(.Data, MSDOSHeaderE.ImageSizeModulo):X} bytes.{NewLine}")
            NewText.Append($"Image size in pages (of 0x200 bytes): 0x{GET_WORD(.Data, MSDOSHeaderE.ImageSize):X}.{NewLine}")
            NewText.Append($"Number of relocation items: 0x{GET_WORD(.Data, MSDOSHeaderE.RelocationCount):X}.{NewLine}")
            NewText.Append($"Header sizes in paragraphs (of 0x10 bytes): 0x{GET_WORD(.Data, MSDOSHeaderE.HeaderSize):X}.{NewLine}")
            NewText.Append($"Minimum number of paragraphs required (of 0x10 bytes): 0x{GET_WORD(.Data, MSDOSHeaderE.MinimumParagraphs):X}.{ NewLine}")
            NewText.Append($"Maximum number of paragraphs required (of 0x10 bytes): 0x{GET_WORD(.Data, MSDOSHeaderE.MaximumParagraphs):X}.{NewLine}")
            NewText.Append($"Stack segment (SS) register: 0x{GET_WORD(.Data, MSDOSHeaderE.StackSegment):X}.{NewLine}")
            NewText.Append($"Stack pointer (SP) register: 0x{GET_WORD(.Data, MSDOSHeaderE.StackPointer):X}.{NewLine}")
            NewText.Append($"Negative checksum of PGM: 0x{ GET_WORD(.Data, MSDOSHeaderE.Checksum):X}.{NewLine}")
            NewText.Append($"Instruction pointer (IP) register: 0x{GET_WORD(.Data, MSDOSHeaderE.InstructionPointer):X}.{NewLine}")
            NewText.Append($"Code segment (CS) register: 0x{GET_WORD(.Data, MSDOSHeaderE.CodeSegment):X}.{ NewLine}")
            NewText.Append($"Relocation table offset: 0x{RelocationTableOffset:X}.{NewLine}")
            NewText.Append($"Overlay number: 0x{GET_WORD(.Data, MSDOSHeaderE.OverlayNumber):X}.{NewLine}")
            NewText.Append(NewLine)
            NewText.Append($"Relocation items:{NewLine}")

            For Position As Integer = RelocationTableOffset To RelocationTableOffset + (RelocationCount * &H4%) Step &H4%
               NewText.Append($"{GET_DWORD(.Data, Position):X8}{NewLine}")
            Next Position

            UpdateDataBox(NewText.ToString())
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the current executable's information palettes.
   Private Sub DisplayPalettesMenu_Click(sender As Object, e As EventArgs) Handles DisplayPalettesMenu.Click
      Try
         Dim Descriptions As New List(Of String)

         For Each location As Integer In CType([Enum].GetValues(GetType(PaletteLocationsE)), Integer())
            Descriptions.Add(DirectCast(location, PaletteLocationsE).ToString())
         Next location

         UpdateDataBox(GBRToText("The executable's palettes:", Palettes(), Descriptions))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure exports the executable's images.
   Public Overloads Sub Export(ExportPath As String)
      Try
         Dim BytesPerRow As New Integer
         Dim IconHeight As New Integer
         Dim IconSize As New Integer
         Dim IconWidth As New Integer

         ExportPath = Path.Combine(ExportPath, Path.GetFileNameWithoutExtension("Cartooners Executable"))
         Directory.CreateDirectory(ExportPath)

         File.WriteAllBytes(Path.Combine(ExportPath, "Cartoons.Unpacked.exe"), DataFile().Data.ToArray())

         Draw4BitImage(DecompressRLE(DataFile().Data, ImageLocationsE.MenuScreen, IMAGE_SIZES(ImageLocationsE.MenuScreen)), SCREEN_WIDTH, SCREEN_HEIGHT, Palettes()(PalettesE.MenuScreenEGA), BYTES_PER_ROW).Save(Path.Combine(ExportPath, "Menu screen.png"), ImageFormat.Png)
         Draw4BitImage(DecompressRLE(DataFile().Data, ImageLocationsE.TicketScreen, IMAGE_SIZES(ImageLocationsE.TicketScreen)), SCREEN_WIDTH, SCREEN_HEIGHT, Palettes()(PalettesE.TicketScreenEGA), BYTES_PER_ROW).Save(Path.Combine(ExportPath, "Ticket screen.png"), ImageFormat.Png)
         Draw4BitImage(DecompressRLE(DataFile().Data, ImageLocationsE.TitleScreen, IMAGE_SIZES(ImageLocationsE.TitleScreen)), SCREEN_WIDTH, SCREEN_HEIGHT, Palettes()(PalettesE.TitleScreenEGA), BYTES_PER_ROW).Save(Path.Combine(ExportPath, "Title screen.png"), ImageFormat.Png)

         For Each Location As Integer In CType([Enum].GetValues(GetType(IconLocationsE)), Integer())
            IconHeight = GET_WORD(DataFile().Data, Location + &H2%)
            IconWidth = GET_WORD(DataFile().Data, Location + &H4%)
            BytesPerRow = If(IconWidth Mod PIXELS_PER_BYTE = 0, IconWidth \ PIXELS_PER_BYTE, (IconWidth + &H1%) \ PIXELS_PER_BYTE)
            IconSize = BytesPerRow * IconHeight
            Draw4BitImage(GetBytes(DataFile().Data, Location + &H6%, IconSize), IconWidth, IconHeight, Palettes()(PalettesE.TitleScreenEGA), BytesPerRow).Save($"{Path.Combine(ExportPath, DirectCast(Location, IconLocationsE).ToString())}.png", ImageFormat.Png)
         Next Location

         Process.Start(New ProcessStartInfo With {.FileName = ExportPath, .WindowStyle = ProcessWindowStyle.Normal})
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the executable's palettes.
   Private Function Palettes(Optional Refresh As Boolean = False) As List(Of List(Of Color))
      Try
         Static CurrentPalettes As New List(Of List(Of Color))

         If Refresh Then
            CurrentPalettes.Clear()

            For Each Location As Integer In CType([Enum].GetValues(GetType(PaletteLocationsE)), Integer())
               CurrentPalettes.Add(GBRPalette(DataFile().Data, Location))
            Next Location
         End If

         Return CurrentPalettes
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Class
