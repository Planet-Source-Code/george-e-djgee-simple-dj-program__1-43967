Attribute VB_Name = "mWinAPI"
Option Explicit
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Const MAX_PATH = 260
Private Const AMAX_PATH = 260

Public Enum FILE_ATTRIBUTES
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * AMAX_PATH
End Type
Public Type VBRinfo
    VBRrate As String
    VBRlength As String
End Type
Public Type MP3Info
    BITRATE As String
    CHANNELS As String
    COPYRIGHT As String
    CRC As String
    EMPHASIS As String
    FREQ As String
    LAYER As String
    LENGTH As String
    MPEG As String
    ORIGINAL As String
    SIZE As String
End Type

Public MP3FileName As String
Public sReturnBuffer As String * 30
Public wichMp3 As Integer
Public MP3Size As Long

Private MP3Length As Long
Private MP3File As String
Private m_scomputername As String

Public Function Between(ByVal accNum As Byte, ByVal accDown As Byte, ByVal accUp As Byte) As Boolean

    If accNum >= accDown And accNum <= accUp Then
        Between = True
      Else 'NOT ACCNUM...
        Between = False
    End If

End Function

Public Sub CloseMp3(Index As Integer)

    mciSendString "close mp3" & Index, 0, 0, 0

End Sub

Private Function Deljivo(ByVal Num As Long) As Byte

    If Num Mod 3 = 0 Then
        Deljivo = 1
      Else 'NOT NUM...
        Deljivo = 0
    End If

End Function

Public Function ftnReturnNodePath(sExplorerPath As String) As String

  Dim iSearch(1) As Integer
  Dim sRootPath As String

    iSearch%(0) = InStr(1, sExplorerPath$, "(", vbTextCompare)
    iSearch%(1) = InStr(1, sExplorerPath$, ")", vbTextCompare)
    If iSearch%(0) > 0 Then
        sRootPath$ = Mid(sExplorerPath$, iSearch%(0) + 1, 2)
    End If
    If iSearch%(1) > 0 Then
        ftnReturnNodePath$ = sRootPath$ & Mid(sExplorerPath$, iSearch%(1) + 1, Len(sExplorerPath$)) & "\"
    End If

End Function

Public Function ftnStripNullChar(sInput As String) As String

  Dim x As Integer

    x = InStr(1, sInput$, Chr$(0))
    If x > 0 Then
        ftnStripNullChar = Left(sInput$, x - 1)
    End If

End Function

Public Sub getMP3Info(ByVal lpMP3File As String, ByRef lpMP3Info As MP3Info)

  Dim Buf As String * 4096
  Dim infoStr As String * 3
  Dim lpVBRinfo As VBRinfo
  Dim tmpByte As Byte
  Dim tmpNum As Byte
  Dim i As Integer
  Dim designator As Byte
  Dim baseFreq As Single
  Dim vbrBytes As Long

    Open lpMP3File For Binary As #1
    Get #1, 1, Buf
    Close #1

    For i = 1 To 4092
        If Asc(Mid(Buf, i, 1)) = &HFF Then
            tmpByte = Asc(Mid(Buf, i + 1, 1))
            If Between(tmpByte, &HF2, &HF7) Or Between(tmpByte, &HFA, &HFF) Then
                Exit For '>---> Next
            End If
        End If
    Next i
    If i = 4093 Then
        MsgBox "Not a MP3 file...", vbCritical, "Error..."
      Else 'NOT I...
        infoStr = Mid(Buf, i + 1, 3)
        'Getting info from 2nd byte(MPEG,Layer type and CRC)
        tmpByte = Asc(Mid(infoStr, 1, 1))

        'Getting CRC info
        If ((tmpByte Mod 16) Mod 2) = 0 Then
            lpMP3Info.CRC = "Yes"
          Else 'NOT ((TMPBYTE...
            lpMP3Info.CRC = "No"
        End If

        'Getting MPEG type info
        If Between(tmpByte, &HF2, &HF7) Then
            lpMP3Info.MPEG = "MPEG 2.0"
            designator = 1
          Else 'NOT BETWEEN(TMPBYTE,...
            lpMP3Info.MPEG = "MPEG 1.0"
            designator = 2
        End If

        'Getting layer info
        If Between(tmpByte, &HF2, &HF3) Or Between(tmpByte, &HFA, &HFB) Then
            lpMP3Info.LAYER = "layer 3"
          Else 'NOT BETWEEN(TMPBYTE,...
            If Between(tmpByte, &HF4, &HF5) Or Between(tmpByte, &HFC, &HFD) Then
                lpMP3Info.LAYER = "layer 2"
              Else 'NOT BETWEEN(TMPBYTE,...
                lpMP3Info.LAYER = "layer 1"
            End If
        End If

        'Getting info from 3rd byte(Frequency, Bit-rate)
        tmpByte = Asc(Mid(infoStr, 2, 1))

        'Getting frequency info
        If Between(tmpByte Mod 16, &H0, &H3) Then
            baseFreq = 22.05
          Else 'NOT BETWEEN(TMPBYTE...
            If Between(tmpByte Mod 16, &H4, &H7) Then
                baseFreq = 24
              Else 'NOT BETWEEN(TMPBYTE...
                baseFreq = 16
            End If
        End If
        lpMP3Info.FREQ = baseFreq * designator * 1000 & " Hz"

        'Getting Bit-rate
        tmpNum = tmpByte \ 16 Mod 16
        If designator = 1 Then
            If tmpNum < &H8 Then
                lpMP3Info.BITRATE = tmpNum * 8
              Else 'NOT TMPNUM...
                lpMP3Info.BITRATE = 64 + (tmpNum - 8) * 16
            End If
          Else 'NOT DESIGNATOR...
            If tmpNum <= &H5 Then
                lpMP3Info.BITRATE = (tmpNum + 3) * 8
              Else 'NOT TMPNUM...
                If tmpNum <= &H9 Then
                    lpMP3Info.BITRATE = 64 + (tmpNum - 5) * 16
                  Else 'NOT TMPNUM...
                    If tmpNum <= &HD Then
                        lpMP3Info.BITRATE = 128 + (tmpNum - 9) * 32
                      Else 'NOT TMPNUM...
                        lpMP3Info.BITRATE = 320
                    End If
                End If
            End If
        End If
        MP3Length = FileLen(lpMP3File) \ (val(lpMP3Info.BITRATE) / 8) \ 1000
        If Mid(Buf, i + 36, 4) = "Xing" Then
            vbrBytes = Asc(Mid(Buf, i + 45, 1)) * &H10000
            vbrBytes = vbrBytes + (Asc(Mid(Buf, i + 46, 1)) * &H100&)
            vbrBytes = vbrBytes + Asc(Mid(Buf, i + 47, 1))
            GetVBRrate lpMP3File, vbrBytes, lpVBRinfo
            lpMP3Info.BITRATE = lpVBRinfo.VBRrate
            lpMP3Info.LENGTH = lpVBRinfo.VBRlength
          Else 'NOT MID(BUF,...
            lpMP3Info.BITRATE = lpMP3Info.BITRATE & "Kbit"
            lpMP3Info.LENGTH = MP3Length & " seconds"
        End If

        'Getting info from 4th byte(Original, Emphasis, Copyright, Channels)
        tmpByte = Asc(Mid(infoStr, 3, 1))
        tmpNum = tmpByte Mod 16

        'Getting Copyright bit
        If tmpNum \ 8 = 1 Then
            lpMP3Info.COPYRIGHT = " Yes"
            tmpNum = tmpNum - 8
          Else 'NOT TMPNUM...
            lpMP3Info.COPYRIGHT = " No"
        End If

        'Getting Original bit
        If (tmpNum \ 4) Mod 2 Then
            lpMP3Info.ORIGINAL = " Yes"
            tmpNum = tmpNum - 4
          Else 'NOT (TMPNUM...
            lpMP3Info.ORIGINAL = " No"
        End If

        'Getting Emphasis bit
        Select Case tmpNum
          Case 0
            lpMP3Info.EMPHASIS = " None"
          Case 1
            lpMP3Info.EMPHASIS = " 50/15 microsec"
          Case 2
            lpMP3Info.EMPHASIS = " invalid"
          Case 3
            lpMP3Info.EMPHASIS = " CITT j. 17"
        End Select

        'Getting channel info
        tmpNum = (tmpByte \ 16) \ 4
        Select Case tmpNum
          Case 0
            lpMP3Info.CHANNELS = " Stereo"
          Case 1
            lpMP3Info.CHANNELS = " Joint Stereo"
          Case 2
            lpMP3Info.CHANNELS = " 2 Channel"
          Case 3
            lpMP3Info.CHANNELS = " Mono"
        End Select
    End If
    lpMP3Info.SIZE = FileLen(lpMP3File) & " bytes"

End Sub

Private Sub GetVBRrate(ByVal lpMP3File As String, ByVal byteRead As Long, ByRef lpVBRinfo As VBRinfo)

  Dim i As Long
  Dim ok As Boolean

    i = 0
    byteRead = byteRead - &H39
    Do
        If byteRead > 0 Then
            i = i + 1
            byteRead = byteRead - 38 - Deljivo(i)
          Else 'NOT BYTEREAD...
            ok = True
        End If
    Loop Until ok
    lpVBRinfo.VBRlength = Trim(Str(i)) & " seconds"
    lpVBRinfo.VBRrate = Trim(Str(Int(8 * FileLen(lpMP3File) / (1000 * i)))) & " Kbit (VBR)"

End Sub

Public Sub openMp3(path As String, Index As Integer)

    mciSendString "open " & path & " type MPEGVideo alias mp3" & Index, 0, 0, 0

End Sub

Public Sub PauseMp3(Index As Integer)

    mciSendString "pause mp3" & Index, 0, 0, 0

End Sub

Public Sub PlayMp3(Index As Integer, Optional flag As String)

    mciSendString "play mp3" & Index & " " & flag, 0, 0, 0

End Sub

Property Get sComputerName() As String ':( Missing Scope

    sComputerName = m_scomputername

End Property

Property Let sComputerName(newValue As String) ':( As Variant ?':( Missing Scope

    m_scomputername = newValue

End Property

Public Sub setMp3backwards(Index As Integer, currTime As Long)

    mciSendString "play mp3" & Index & " reverse", 0, 0, 0

End Sub

Public Sub setMp3Bass(Index As Integer, newValue As Integer)

    mciSendString "setaudio mp3" & Index & " bass to " & newValue, 0, 0, 0

End Sub

Public Sub setMP3channelState(Index As Integer, cHannel As String, sTate As String)

  'off or on for state

    mciSendString "set mp3" & Index & " audio " & cHannel & " " & sTate, 0, 0, 0

End Sub

Public Sub setMp3channelVolume(Index As Integer, cHannel As String, newValue As Integer)

    mciSendString "setaudio mp3" & Index & " " & cHannel & " volume to " & newValue, 0, 0, 0

End Sub

Public Sub setMp3CurrentTime(Index As Integer, newTime As Double, flags As String)

    mciSendString "play mp3" & Index & " from " & newTime & " " & flags, 0, 0, 0

End Sub

Public Sub setMp3Speed(Index As Integer, speed As Integer)

    mciSendString "set mp3" & Index & " speed" & " " & speed, 0, 0, 0

End Sub

Public Sub setMp3State(sTate As String, Index As Integer)

    mciSendString "set " & sTate & " mp3" & Index, 0, 0, 0

End Sub

Public Sub setMp3TimeFormat(Index As Integer, TimeFormat As String)

    mciSendString "set mp3" & Index & " time format " & TimeFormat, 0, 0, 0

End Sub

Public Sub setMp3Treble(Index As Integer, newValue As Integer)

    mciSendString "setaudio mp3" & Index & " treble to " & newValue, 0, 0, 0

End Sub

Public Sub setMp3Volume(Index As Integer, newValue As Integer)

    mciSendString "setaudio mp3" & Index & " volume to " & newValue, 0, 0, 0

End Sub

Public Function statusBass(ByVal Index As Integer) As Integer

    mciSendString "status  mp3" & Index & " bass", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusBass = val(sReturnBuffer$)

End Function

Public Function statusChannelVolume(ByVal Index As Integer, ByVal cHannel As String) As Integer

    mciSendString "status mp3" & Index & " " & cHannel & " volume", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusChannelVolume = val(sReturnBuffer$)

End Function

Public Function statusLength(ByVal Index As Integer) As Long

    mciSendString "status mp3" & Index & " length", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusLength = val(sReturnBuffer$)

End Function

Public Function statusMp3state(ByVal Index As Integer) As String

    mciSendString "status mp3" & Index & " audio", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusMp3state = Left(sReturnBuffer$, 2)

End Function

Public Function statusPosition(ByVal Index As Integer) As Double

    mciSendString "status mp3" & Index & " position", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusPosition = val(sReturnBuffer$)

End Function

Public Function statusSpeed(ByVal Index As Integer) As Integer

    mciSendString "status mp3" & Index & " speed", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusSpeed = val(sReturnBuffer$)

End Function

Public Function statusTreble(ByVal Index As Integer) As Integer

    mciSendString "status  mp3" & Index & " treble", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusTreble = val(sReturnBuffer$)

End Function

Public Function statusVolume(ByVal Index As Integer) As Integer

    mciSendString "status mp3" & Index & " volume", sReturnBuffer$, Len(sReturnBuffer$), 0
    statusVolume = val(sReturnBuffer$)

End Function

Public Sub StopMp3(Index As Integer)

    mciSendString "stop mp3" & Index, 0, 0, 0

End Sub

Public Sub subFileList(sFolderPath As String)

  Dim lReturn As Long                    'Search Handle of specified Path.
  Dim lNextFile As Long                  'Search Handle of specified File.
  Dim sPath As String                    'Path to search.
  Dim WFD As WIN32_FIND_DATA             'Set Variable WFD as Structure(VBType) WIN32_FIND_DATA.
  Dim lstItem As ListItem                'lstItem = A ListView ListItem.
  Dim lstSubItem As ListSubItem          'lstSubItem = A ListView ListSubItem.
  Dim sFileName As String                'Filename (WFD.cFileName).
  Dim oFileList As ListView              'Set oFileList as Control being used.

    Set oFileList = frmExplore.FileList
    sPath$ = sFolderPath$ & "*.mp3"
  Dim lFileLoop As Long                  'Loop for setting ForeColour of specific Files. eg(*.exe).':( Move line to top of current Sub

    With oFileList

        .Visible = False
        .ListItems.Clear

        lReturn& = FindFirstFile(sPath$, WFD) & Chr$(0)
        frmExplore.MousePointer = 11

        Do

            'If we find a Directory do nothing, else List Files taking off the Chr$(0)
            'Loop until lNextFile& = val(0), no more Files to List
            If Not (WFD.dwFileAttributes And vbDirectory) = vbDirectory Then

                sFileName$ = ftnStripNullChar(WFD.cFileName)

                If sFileName > Trim("") Then
                    Set lstItem = .ListItems.Add(, , sFileName$)
                    Set lstSubItem = lstItem.ListSubItems.Add(, , Format(WFD.nFileSizeLow, "#,0"))
                End If
            End If
            lNextFile& = FindNextFile(lReturn&, WFD)
        Loop Until lNextFile& <= val(0)
        frmExplore.MousePointer = 0
        lNextFile& = FindClose(lReturn&)
        For lFileLoop = 1 To .ListItems.Count
            If InStrRev(LCase(.ListItems(lFileLoop).Text), ".mp3", , vbTextCompare) Then
                .ListItems(lFileLoop).ForeColor = RGB(60, 60, 140)
            End If
        Next lFileLoop
        .Visible = True
    End With 'OFILELIST

End Sub

Public Sub subShowFolderList(oFolderList As ListBox, oExplorerTree As TreeView, sDriveLetter As String, vParentID As Variant)

  Dim nNode As Node                           'Node object for DirTree.
  Dim lReturn As Long                         'Holds Search Handle of File.
  Dim lNextFile As Long                       'Return Search Handle of next Folder.
  Dim sPath As String                         'Path to search.
  Dim WFD As WIN32_FIND_DATA                  'Win32 Structure (VB Type).
  Dim sFolderName As String                   'Name of Folder.
  Dim x As Long                               'Used to loop through Folders in frmMain.List1).

    Set oFolderList = frmExplore.List1          'Set Object oFolderList as frmMain.List1.
    Set oExplorerTree = frmExplore.Explorer     'Set Object oExplorerTree as source Explorer Tree.
    sPath$ = (sDriveLetter & "*.*") & Chr$(0)
    lReturn& = FindFirstFile(sPath$, WFD)
    Do
        If (WFD.dwFileAttributes And vbDirectory) Then
            sFolderName$ = ftnStripNullChar(WFD.cFileName)
            If sFolderName$ <> "." And sFolderName$ <> ".." Then
                If WFD.dwFileAttributes <> 16 Then
                    oFolderList.AddItem sFolderName$ & "~A~"
                  Else 'NOT WFD.DWFILEATTRIBUTES...
                    oFolderList.AddItem sFolderName$ & "~~~"
                End If
            End If
        End If
        lNextFile& = FindNextFile(lReturn&, WFD)
    Loop Until lNextFile& = False
    lNextFile& = FindClose(lReturn&)
    For x = 0 To oFolderList.ListCount - 1
        If Right(oFolderList.List(x), 3) = "~A~" Then
            Set nNode = oExplorerTree.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
            nNode.ForeColor = RGB(120, 120, 120)
          Else 'NOT RIGHT(OFOLDERLIST.LIST(X),...
            Set nNode = oExplorerTree.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
        End If
    Next x
    oFolderList.Clear

End Sub

Public Sub subLoadTreeView()

    Dim TreeList As TreeView                    'Explorer Tree.
        Set TreeList = frmExplore.Explorer
    Dim iDriveNum As Integer                    'Key Index in DirTree.
    Dim sDriveType As String                    'Holds DriveType.
    Dim fso, d As Object                        'Used to return DriveType.
        Set fso = CreateObject("Scripting.FileSystemObject")
    Dim x As Integer                            'Loop through Drives.
    Dim RetStr(1) As String                     'Holds Drive letters.
    Dim nNode As Node                           'Node object for ExplorerTree.
    Dim sComputerName As String                 'Hold Computer Name.
        sComputerName$ = mWinAPI.sComputerName
                        
    'Return Drive structure from XFile.Dll.-----------------------------------------
    RetStr$(0) = ftnShowDriveList
    '-------------------------------------------------------------------------------

    With TreeList

        'Add Computer name to DirTree-----------------------------------------------
        Set nNode = .Nodes.Add(, , sComputerName$, sComputerName$, "mycomputer", "mycomputer")
        '---------------------------------------------------------------------------
        
        'Add Drive A:\ to DirTree---------------------------------------------------
'        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & "0", "3.5 Floppy (A:)", "drvremove") 'Add Drive "A:"
        '---------------------------------------------------------------------------
        
        'Loop through RetStr$(0) to retrieve Drives. eg."ACDEF".--------------------
        For x = 1 To Len(RetStr$(0))
            
            'Strip Driveinfo eg "A"-------------------------------------------------
            RetStr$(1) = Mid(RetStr$(0), x, 1)
            '-----------------------------------------------------------------------
            
            'Get DriveType information.---------------------------------------------
            Set d = fso.GetDrive(RetStr$(1))
            '-----------------------------------------------------------------------
            
            'Used to make unique Key Index in DirTree.------------------------------
            iDriveNum% = x
            '-----------------------------------------------------------------------
                       
            'Determine Drive type and add to TreeList.------------------------------
            Select Case d.drivetype
                
                'Unknown Drive.-----------------------------------------------------
                Case 0: sDriveType = "Unknown"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvunknown")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvunknown")
                    End If

                'Removable Drive.---------------------------------------------------
                Case 1: sDriveType = "Removable"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":", "drvremove")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvremove")

                    End If
                    
                'Fixed Drive.-------------------------------------------------------
                Case 2: sDriveType = "Fixed"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvfixed")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvfixed")
                    End If

                'Network Drive.-----------------------------------------------------
                Case 3: sDriveType = "Network"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvremote")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvremote")
                    End If
                    
                'CD-Rom.------------------------------------------------------------
                Case 4: sDriveType = "CD-ROM"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvcd")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvcd")
                    End If
                    
                'Ram Disk.----------------------------------------------------------
                Case 5: sDriveType = "Ram Disk"
                    If d.isready Then
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, d.volumename & " (" & d.driveletter & ":)", "drvram")
                    Else
                        Set nNode = .Nodes.Add(sComputerName$, tvwChild, "Parent" & iDriveNum%, " (" & d.driveletter & ":)", "drvram")
                    End If
            
            End Select
            '-----------------------------------------------------------------------
            
        Next x

    End With

End Sub

Private Function ftnShowDriveList()
  
    Dim fso, d, dc As Object
    Dim sDriveLetter As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set dc = fso.Drives
    
    For Each d In dc
        sDriveLetter$ = sDriveLetter$ & d.driveletter
    Next
    
    ftnShowDriveList = sDriveLetter$

End Function

