'This class' imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

'This class contains the movie related procedures.
Public Class MovieClass
   Inherits DataFileClass

   'This enumeration contains the locations of known values inside a movie file.
   Private Enum LocationsE As Integer
      ActorsScenesPalette = &H4%          'The movie's first palette.
      DefaultSpeechBalloonColor = &HE1%   'The movie's default speech balloon text color palette index.
      FrameCount = &HE3%                  'The movie's frame count.
      FrameData = &HE5%                   'The movie's frame data.
      FrameRate = &H26%                   'The movie's frame rate.
      HasScenes = &H24%                   'Indicates whether or not the movie contains scenes.
      HasSpeechBalloons = &HDF%           'Indicates whether or not the movie contains speech balloons.
      PlayRepeatedly = &H28%              'Indicates whether or not the movie is played repeatedly.
      Signature = &H0%                    'The movie file signature.
      SpeechBalloonsPalette = &HBF%       'The movie's second palette.
   End Enum

   'This enumeration contains a list of the palettes inside a movie file.
   Private Enum PalettesE As Integer
      ActorsScenesPalette            'The movie's actors and scenes palette.
      SpeechBalloonsPalette          'The movie's speech balloons palette.
   End Enum

   'This enumeration contains the locations of known values inside a speech balloon header.
   Private Enum SpeechBalloonLocationsE As Integer
      Alignment = &H24%                  'The text's alignment.
      BackColor = &H6%                   'The balloon background color's palette index.
      BalloonHeight = &HE%               'The balloon's height.
      BalloonWidth = &H10%               'The balloon's width.
      BorderColor = &H8%                 'The balloon border color's palette index.
      Index = &H0%                       'The balloon's index.
      LastSelectedTextColor = &H20%      'The last selected text color's palette index.
      PropertiesSize = &H2%              'The balloon properties' size.
      Text = &H2A%                       'The balloon's text.
      TextHeight = &H16%                 'The balloon's text area height.
      TextWidth = &H18%                  'The balloon's text area width.
      TextX = &H12%                      'The text's vertical position.
      TextY = &H14%                      'The text's horizontal position.
      TextLength = &H1A%                 'The text's length.
      Type = &H4%                        'The balloon's type.
   End Enum

   'This structure defines a movie's speech balloon.
   Private Structure SpeechBalloonStr
      Public Header As List(Of Byte)     'The balloon's header.
      Public Text As String              'The balloon's text.
   End Structure

   'The movie related constants used by this program.
   Private Const MAXIMUM_FRAME_RATE As Integer = 60                      'Contains the highest number of frames per second supported.
   Private Const MINIMUM_INTERVAL As Double = 1000 / MAXIMUM_FRAME_RATE  'Contains the lowest number of milliseconds between frames supported.
   Private Const UNKNOWN_DWORDS_PER_ACTOR As Integer = &H5%              'Contains the number of unknown DWORDs per actor.
   Private ReadOnly PALETTE_DESCRIPTIONS As New List(Of String)({"actors and scenes", "speech balloons"})                                                                                                     'Contains the movie palettes descriptions.
   Private ReadOnly SIGNATURE As New List(Of Byte)({&H10%, &H10%, &HDF%, &H0%})                                                                                                                               'Contains the movie file signature.
   Private ReadOnly SPEECH_BALLOON_ALIGNMENTS As New List(Of String)({"left", "center"})                                                                                                                      'Contains the movie speech balloon alignments.
   Private ReadOnly SPEECH_BALLOON_TYPES As New List(Of String)({"Invisible", "Title", "Speech (Right)", "Speech (Left)", "Thought (Right)", "Thought (Left)", "Exclamation (Right)", "Exclamtion (Left)"})   'Contains the movie's speech balloon types.

   'The menu items used by this class.
   Private WithEvents DisplayFilesMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F1, .Text = "Display &Files"}
   Private WithEvents DisplayFrameRecordsMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F2, .Text = "Display Frame &Records"}
   Private WithEvents DisplayInformationMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F3, .Text = "Display &Information"}
   Private WithEvents DisplayPalettesMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F4, .Text = "Display &Palettes"}
   Private WithEvents DisplaySpeechBalloonsMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F5, .Text = "Display &Speech Balloons"}
   Private WithEvents DisplayUnknownDataMenu As New ToolStripMenuItem With {.ShortcutKeys = Keys.F6, .Text = "Display &Unknown Data"}

   'This procedure initializes this class.
   Public Sub New(PathO As String, Optional DataFileMenu As ToolStripMenuItem = Nothing)
      Try
         If DataFile(MoviePath:=PathO).Data.Count > 0 AndAlso DataFileMenu IsNot Nothing Then
            With DataFileMenu
               .DropDownItems.Clear()
               .DropDownItems.AddRange({DisplayFilesMenu, DisplayFrameRecordsMenu, DisplayInformationMenu, DisplayPalettesMenu, DisplaySpeechBalloonsMenu, DisplayUnknownDataMenu})
               .Text = "&Movie"
               .Visible = True
            End With
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure returns the movie actors' handles.
   Private Function ActorHandles(Optional Refresh As Boolean = False, Optional ActorCount As Integer = Nothing, Optional ByRef Position As Integer = Nothing) As List(Of Integer)
      Try
         Static CurrentActorHandles As New List(Of Integer)

         If Refresh AndAlso Not Position = Nothing Then
            CurrentActorHandles.Clear()

            Position += &H1%
            For Actor As Integer = &H1% To ActorCount - &H1%
               CurrentActorHandles.Add(DataFile().Data(Position) + &H1%)
               Position += &H1%
            Next Actor
            Position += &H3%
         End If

         Return CurrentActorHandles
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedures manages the movie's data.
   Private Function DataFile(Optional MoviePath As String = Nothing) As DataFileStr
      Try
         Dim Position As Integer = 0
         Static CurrentFile As New DataFileStr With {.Data = Nothing, .Path = Nothing}

         If Not MoviePath = Nothing Then
            With CurrentFile
               .Data = New List(Of Byte)(File.ReadAllBytes(MoviePath))

               If GetBytes(CurrentFile.Data, LocationsE.Signature, SIGNATURE.Count).SequenceEqual(SIGNATURE) Then
                  .Path = MoviePath
                  FrameRecords(Refresh:=True, Position:=Position)
                  UnknownDWords(Refresh:=True, Position:=Position)
                  Files(Refresh:=True, Position:=Position)
                  SpeechBalloons(Refresh:=True, Position:=Position)
                  Palettes(Refresh:=True)

                  DisplayInformationMenu.PerformClick()
               Else
                  MessageBox.Show("Invalid movie file.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                  .Data = Nothing
               End If
            End With
         End If

         Return CurrentFile
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure displays the movie's list of external files.
   Private Sub DisplayFilesMenu_Click(sender As Object, e As EventArgs) Handles DisplayFilesMenu.Click
      Try
         With New StringBuilder
            .Append($"Movie files:{NewLine}")

            For Index As Integer = 0 To Files().Count - 1
               .Append($"""{Files()(Index)}""")
               If Index < ActorHandles().Count Then .Append($" ({ActorHandles()(Index):X})")
               .Append(NewLine)
            Next Index

            UpdateDataBox(.ToString())
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the movie's frame data.
   Private Sub DisplayFrameRecordsMenu_Click(sender As Object, e As EventArgs) Handles DisplayFrameRecordsMenu.Click
      Try
         Dim Length As New Integer

         With New StringBuilder
            .Append($"Frame record data:{NewLine}")
            For Record As Integer = 0 To FrameRecords().Count - 1
               Length = FrameRecords()(Record).Length
               .Append($"{NewLine}Frame Record #{Record} - Length: {Length}{NewLine}")
               .Append(Escape(GetString(FrameRecords()(Record).ToList(), &H0%, Length), " "c, EscapeAll:=True).Trim())
               .Append(NewLine)
            Next Record

            UpdateDataBox(.ToString())
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the general information for the current movie.
   Private Sub DisplayInformationMenu_Click(sender As Object, e As EventArgs) Handles DisplayInformationMenu.Click
      Try
         With New StringBuilder
            .Append($"General information:{NewLine}")
            .Append($"-Path: {DataFile().Path}{NewLine}")
            .Append($"-Frames: {BitConverter.ToUint16(DataFile().Data.ToArray(), LocationsE.FrameCount)}{NewLine}")
            .Append($"-Frames per second: {FrameRate()}{NewLine}")
            .Append($"-Play repeatedly: {CBool(BitConverter.ToUint16(DataFile().Data.ToArray(), LocationsE.PlayRepeatedly))}{NewLine}")
            .Append($"-Contains scenes: {CBool(BitConverter.ToUint16(DataFile().Data.ToArray(), LocationsE.HasScenes))}{NewLine}")
            .Append($"-Contains speech balloons: {CBool(BitConverter.ToUint16(DataFile().Data.ToArray(), LocationsE.HasSpeechBalloons))}{NewLine}")
            .Append($"-Default speech balloon text color: {BitConverter.ToUint16(DataFile().Data.ToArray(), LocationsE.DefaultSpeechBalloonColor)}")

            UpdateDataBox(.ToString())
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the current movie's palettes.
   Private Sub DisplayPalettesMenu_Click(sender As Object, e As EventArgs) Handles DisplayPalettesMenu.Click
      Try
         UpdateDataBox(GBRToText("The movie's palettes:", Palettes(), PALETTE_DESCRIPTIONS))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the current movie's speech balloons.
   Private Sub DisplaySpeechBalloonsMenu_Click(sender As Object, e As EventArgs) Handles DisplaySpeechBalloonsMenu.Click
      Try
         Dim NewText As New StringBuilder

         NewText.Append($"Speech balloons:{NewLine}{NewLine}")

         For Each SpeechBalloon As SpeechBalloonStr In SpeechBalloons()
            With SpeechBalloon
               NewText.Append($"[Index: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.Index):X}]{NewLine}")
               NewText.Append($"Alignment: {SPEECH_BALLOON_ALIGNMENTS(BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.Alignment))}{NewLine}")
               NewText.Append($"Type: {SPEECH_BALLOON_TYPES(BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.Type))}{NewLine}")
               NewText.Append($"Background color: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.BackColor):X}{NewLine}")
               NewText.Append($"Balloon width: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.BalloonWidth):X}{NewLine}")
               NewText.Append($"Balloon height: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.BalloonHeight):X}{NewLine}")
               NewText.Append($"Border color: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.BorderColor):X}{NewLine}")
               NewText.Append($"Last selected text color: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.LastSelectedTextColor):X}{NewLine}")
               NewText.Append($"Text area width: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.TextWidth):X}{NewLine}")
               NewText.Append($"Text area height: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.TextHeight):X}{NewLine}")
               NewText.Append($"Text X: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.TextX):X}{NewLine}")
               NewText.Append($"Text Y: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.TextY):X}{NewLine}")
               NewText.Append($"Text length: {BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.TextLength):X}{NewLine}")
               NewText.Append($"Text: ""{Escape(.Text)}""{NewLine}{NewLine}")
            End With
         Next SpeechBalloon

         UpdateDataBox(NewText.ToString())
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the current movie's unknown data.
   Private Sub DisplayUnknownDataMenu_Click(sender As Object, e As EventArgs) Handles DisplayUnknownDataMenu.Click
      Try
         Dim DWORDText As String = Nothing

         With New StringBuilder
            .Append($"Unknown data:{NewLine}{NewLine}")
            .Append($"{"Actor:",-15}")
            For Label As Integer = 1 To UNKNOWN_DWORDS_PER_ACTOR
               .Append($"{$"DWORD {Label:X}:",-17}")
            Next Label

            For Unknown As Integer = 0 To UnknownDWords().Count - 1 Step UNKNOWN_DWORDS_PER_ACTOR
               .Append(NewLine)
               .Append($"{$"{$"{((Unknown \ UNKNOWN_DWORDS_PER_ACTOR) + 1):X}",5}",-15}")

               For DWORD As Integer = Unknown To Unknown + (UNKNOWN_DWORDS_PER_ACTOR - 1)
                  DWORDText = $"{UnknownDWords()(DWORD):X8}"
                  For ByteO As Integer = &H0% To DWORDText.Length - &H1% Step &H2%
                     .Append($"{DWORDText.Substring(ByteO, &H2%)} ")
                  Next ByteO
                  .Append(New String(" "c, 5))
               Next DWORD
            Next Unknown

            UpdateDataBox(.ToString())
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the movie's list of external files.
   Private Function Files(Optional Refresh As Boolean = False, Optional ByRef Position As Integer = Nothing) As List(Of String)
      Try
         Dim FileCount As New Integer
         Dim Item As String = Nothing
         Dim Length As New Integer
         Static CurrentFiles As New List(Of String)

         If Refresh AndAlso Not Position = Nothing Then
            CurrentFiles.Clear()

            FileCount = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
            Position += &H2%

            ActorHandles(Refresh:=True, ActorCount:=FileCount, Position:=Position)

            For ActorPath As Integer = &H1% To FileCount - &H1%
               Length = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
               Position += &H2%
               Item = TERMINATE_AT_NULL(GetString(DataFile().Data, Position, Length, AdvanceOffset:=True))
               CurrentFiles.Add(Item)
            Next ActorPath

            FileCount = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
            Position += &H2%

            For ScenePath As Integer = &H1% To FileCount
               Length = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
               Position += &H2%
               Item = TERMINATE_AT_NULL(GetString(DataFile().Data, Position, Length, AdvanceOffset:=True))
               CurrentFiles.Add(Item)
            Next ScenePath

            FileCount = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
            Position += &H2%

            For MusicPath As Integer = &H1% To FileCount
               Length = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
               Position += &H2%
               Item = TERMINATE_AT_NULL(GetString(DataFile().Data, Position, Length, AdvanceOffset:=True))
               CurrentFiles.Add(Item)
            Next MusicPath
         End If

         Return CurrentFiles
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the movie's frame rate.
   Private Function FrameRate() As Double
      Try
         Return (1000 / BitConverter.ToUint16(DataFile().Data.ToArray(), LocationsE.FrameRate)) / MINIMUM_INTERVAL
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the movie's frame records.
   Private Function FrameRecords(Optional Refresh As Boolean = False, Optional ByRef Position As Integer = Nothing) As List(Of Byte())
      Try
         Dim FrameCount As Integer = BitConverter.ToUint16(DataFile().Data.ToArray(), LocationsE.FrameCount)
         Dim Length As New Integer
         Static CurrentFrameRecords As New List(Of Byte())

         If Refresh Then
            CurrentFrameRecords.Clear()
            Position = LocationsE.FrameData

            For Record As Integer = &H0% To FrameCount - &H1%
               Length = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
               Position += &H2%
               CurrentFrameRecords.Add(GetBytes(DataFile().Data, Position, Length).ToArray())
               Position += Length
            Next Record
         End If

         Return CurrentFrameRecords
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the movie's palettes.
   Private Function Palettes(Optional Refresh As Boolean = False) As List(Of List(Of Color))
      Try
         Static CurrentPalettes As New List(Of List(Of Color))

         If Refresh Then
            CurrentPalettes.Clear()

            Array.ForEach({LocationsE.ActorsScenesPalette, LocationsE.SpeechBalloonsPalette}, Sub(PaletteLocation As Integer) CurrentPalettes.Add(GBRPalette(DataFile().Data, PaletteLocation)))
         End If

         Return CurrentPalettes
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the movie's speech balloons.
   Private Function SpeechBalloons(Optional Refresh As Boolean = False, Optional ByRef Position As Integer = Nothing) As List(Of SpeechBalloonStr)
      Try
         Dim NewSpeechBalloon As SpeechBalloonStr = Nothing
         Dim PropertiesSize As New Integer
         Dim SpeechBalloonCount As New Integer
         Dim TextLength As New Integer
         Static CurrentSpeechBalloons As List(Of SpeechBalloonStr)

         If Refresh AndAlso Not Position = Nothing Then
            CurrentSpeechBalloons = New List(Of SpeechBalloonStr)
            SpeechBalloonCount = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
            Position += &H2%

            For SpeechBalloon As Integer = 1 To SpeechBalloonCount
               NewSpeechBalloon = New SpeechBalloonStr
               With NewSpeechBalloon
                  PropertiesSize = BitConverter.ToUint16(DataFile().Data.ToArray(), Position + SpeechBalloonLocationsE.PropertiesSize)
                  .Header = New List(Of Byte)(GetBytes(DataFile().Data, Position, PropertiesSize + &H4%, AdvanceOffset:=True))
                  TextLength = BitConverter.ToUint16(.Header.ToArray(), SpeechBalloonLocationsE.TextLength)
                  .Text = GetString(DataFile().Data, Position, TextLength, AdvanceOffset:=True)
               End With
               CurrentSpeechBalloons.Add(NewSpeechBalloon)
            Next SpeechBalloon
         End If

         Return CurrentSpeechBalloons
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the movie's unknown DWords.
   Private Function UnknownDWords(Optional Refresh As Boolean = False, Optional ByRef Position As Integer = Nothing) As List(Of Integer)
      Try
         Dim Count As New Integer
         Static CurrentDWords As New List(Of Integer)

         If Refresh AndAlso Not Position = Nothing Then
            CurrentDWords.Clear()

            Count = BitConverter.ToUint16(DataFile().Data.ToArray(), Position)
            Position += &H2%
            For Index As Integer = &H0% To Count - &H1%
               CurrentDWords.Add(BitConverter.ToInt32(DataFile().Data.ToArray(), Position))
               Position += &H4%
            Next Index
         End If

         Return CurrentDWords
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Class
