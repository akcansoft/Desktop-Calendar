VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "form"
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   6420
   FillStyle       =   0  'Solid
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H0030B4F3&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   3000
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' AS Desktop Calendar
' This application makes the current month
' and next month calendar in local language
' on the desktop image.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' First release: 16/04/2004
' v1.2 : 10/09/2024
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 1) The application can be set to run at Windows startup.
'    To enable this, runAtStartup = True must be set in settings.ini.
' 2) An outline option has been added to the text in the calendar.
'    For this, the following option should be set in settings.ini:
'    textEffect = outline
'    options:
'    none - No effect is applied.
'    shadow - Adds a shadow behind the text.
'    outline- Adds an outline around the text.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Mesut AKCAN
' makcan@gmail.com
' https://akcansoft.blogspot.com
' https://youtube.com/mesutakcan
' https://github.com/akcansoft
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Const wallpaperFname As String = "wallpaper.bmp"
Const holidaysFname As String = "holidays.txt"
Const iniFile As String = "settings.ini"
Const hkcu As String = "HKEY_CURRENT_USER\"
Const regStartUpRunPath As String = hkcu & "Software\Microsoft\Windows\CurrentVersion\Run\"
Const regCpanelDeskopPath As String = hkcu & "Control Panel\Desktop\"
Const COLOR_DESKTOP As Long = 1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = 1
Const SPIF_SENDWININICHANGE = 2
Const HKEY_CURRENT_USER = &H80000001

Private Declare Function RoundRect Lib "gdi32" ( _
  ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
  ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
  ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'Font colors
Dim fontColor As Long 'The color of the calendar font
Dim shadowColor As Long 'The color of the shadow effect
Dim holidayColor As Long 'The color for the holidays text
Dim weekDayColor As Long 'The color for the weekdays text

Dim textEffect As String 'The visual effect applied to the text
Dim currentDayShape As String ' The shape used to highlight the current day.
Dim currentMonth As Byte, currentYear As Integer
Dim weekDays(1 To 7) As String * 2
Dim holidays As String
Dim reg As New Registry

Private Sub Form_Load()
  Dim n As Byte
  Dim activeWallpaper As String
  Dim months(1 To 12) As String
  Dim offsetX As Integer, offsetY As Integer
  Dim posX As Integer, posY As Integer
  Dim prevWallpaper As String, calendarSourceImage As String
  Dim runAtStartup As Boolean, appFullPath As String
  Dim regKey As String, retValue
  
  appFullPath = App.Path & "\" & App.EXEName & ".exe"
  runAtStartup = CBool(ReadINI(iniFile, "APP", "runAtStartup", "False"))
  regKey = regStartUpRunPath & App.ProductName
  retValue = reg.ReadKey(regKey)
  If runAtStartup Then
    If retValue <> appFullPath Then
      retValue = reg.WriteKey(regKey, appFullPath)
    End If
  Else
    If retValue <> "" Then
      retValue = reg.DeleteKey(regKey)
    End If
  End If
  
  shadowColor = CLng(ReadINI(iniFile, "FONT", "shadowColor", "&H000000"))
  fontColor = CLng(ReadINI(iniFile, "FONT", "fontColor", "&HFFFFFF"))
  holidayColor = CLng(ReadINI(iniFile, "FONT", "holidayColor", "&H0000FF"))
  weekDayColor = CLng(ReadINI(iniFile, "FONT", "weekdayColor", "&H15AE4F"))
  textEffect = ReadINI(iniFile, "FONT", "textEffect")
  currentDayShape = ReadINI(iniFile, "SHAPE", "currentDayShape", "RoundRectangle")
  
  ' Read the currently active wallpaper from the registry
  activeWallpaper = reg.ReadKey(regCpanelDeskopPath & "Wallpaper")
  ' Read the previously set wallpaper from the INI file
  prevWallpaper = ReadINI(iniFile, "WALLPAPER", "prevWallpaper")
  
  ' Check if the active wallpaper is the same as the application-generated wallpaper (calendar added)
  If activeWallpaper = App.Path & "\" & wallpaperFname Then
    ' If prevWallpaper exists, set it as the active wallpaper
    If FileExists(prevWallpaper) Then
      activeWallpaper = prevWallpaper
    Else
      ' If prevWallpaper doesn't exist, set activeWallpaper to empty (no wallpaper)
      activeWallpaper = ""
    End If
  End If
  
  ' If the active wallpaper is a valid file
  If FileExists(activeWallpaper) Then
    ' Set calendarSourceImage to the active wallpaper
    calendarSourceImage = activeWallpaper
    ' If the previous wallpaper differs from the active wallpaper, update the INI file
    If prevWallpaper <> activeWallpaper Then
      WriteINI iniFile, "WALLPAPER", "prevWallpaper", activeWallpaper
    End If
  Else
    ' If no valid wallpaper exists, clear calendarSourceImage and delete the prevWallpaper key from the INI file
    calendarSourceImage = ""
    DeleteINIKey iniFile, "WALLPAPER", "prevWallpaper"
  End If
    
  ' Create a list of month names in the local language
  For n = 1 To 12
    months(n) = Format("1/" & n & "/2024", "MMMM") ' Year is irrelevant
  Next
  
  ' Create a list of weekdays in the local language. The first date provided is Monday
  For n = 1 To 7
    weekDays(n) = Format(n & "/7/2024", "ddd") ' 01/07/2024 -> Monday
  Next
  
  ' Load holidays from the file
  holidays = LoadHolidays()
  With Me
    .ScaleMode = vbPixels
    .BorderStyle = 0
    .Appearance = 0
    .Width = Screen.Width
    .Height = Screen.Height
    .Move 0, 0
  End With
  
  currentMonth = Month(Now()) ' Current month
  currentYear = Year(Now()) ' Current year

  With picBox
    .BorderStyle = 0
    '.Appearance = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
    .Move 0, 0
    .Font.Name = ReadINI(iniFile, "FONT", "fontName", "Tahoma")
    .FontBold = CBool(ReadINI(iniFile, "FONT", "fontBold", "True"))
    .FontItalic = CBool(ReadINI(iniFile, "FONT", "fontItalic", "False"))
    
    '~~~~~~~~~~~~ This month settings ~~~~~~~~~~~
    .Font.Size = Me.ScaleHeight / Val(ReadINI(iniFile, "FONT", "fontRatio_1", "45"))
    ' The fill color of the shape for the current day
    .FillColor = CLng(ReadINI(iniFile, "SHAPE", "shapeFillColor", "&H30B4F3"))
    
    ApplyWallpaper calendarSourceImage
    'Set reg = Nothing
    
    ' Position of the calendar start point(top center) + offset
    offsetX = (.Width / 2) - (.TextWidth("8") * 21 / 2) + CInt(ReadINI(iniFile, "CALENDAR POSITION", "startOffsetX", "0")) '21 = 3chr * 7day
    offsetY = CInt(ReadINI(iniFile, "CALENDAR POSITION", "startOffsetY", "10"))
    
    PicPrint offsetX, offsetY, 2, Format(Now, "d MMMM yyyy dddd"), fontColor ' Write today's date
    '~~~~~~~~~~~~ This month calendar ~~~~~~~~~~~
    Calendar offsetX, .CurrentY, 2
    
    '~~~~~~~~~~~~ Next month settings ~~~~~~~~~~~
    .Font.Size = Me.ScaleHeight / Val(ReadINI(iniFile, "FONT", "fontRatio_2", "65"))
    .CurrentX = offsetX
    .CurrentY = .CurrentY + .TextHeight("8") / 2 ' Add half the text height
    currentMonth = currentMonth + 1
    
    ' If the current month is December, the next month will be January of the following year
    If currentMonth > 12 Then
      currentMonth = 1
      currentYear = currentYear + 1
    End If
    
    ' Position of the next month calendar start point
    posX = .CurrentX
    posY = .CurrentY
    
    ' Next month's header
    PicPrint posX, posY, 1, months(currentMonth) & " " & currentYear, fontColor
    
    '~~~~~~~~~~~~ Next month calendar ~~~~~~~~~~~~
    Calendar offsetX, .CurrentY, 1
    
    ' Write akcanSoft
    .CurrentX = posX
    .Font.Size = .Font.Size / 2
    PicPrint posX, .CurrentY, 1, "akcanSoft", &HCCCCCC
  End With
  
  'save picturebox image to file
  Dim wallpaperFileName As String
  wallpaperFileName = App.Path & "\" & wallpaperFname
  SavePicture picBox.Image, wallpaperFileName
  ' Set the saved image file as the desktop wallpaper
  If FileExists(wallpaperFileName) Then
    'reg.WriteKey regCpanelDeskopPath & "Wallpaper", wallpaperFileName
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0, wallpaperFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
    If calendarSourceImage = "" Then SystemParametersInfo SPI_SETDESKWALLPAPER, 0, wallpaperFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
  End If
  'App Exit
  End
End Sub

'~~~~~~~~~~~~ CALENDAR ~~~~~~~~~~~~~~~~~~
Private Sub Calendar(x As Integer, y As Integer, o As Byte)
  Dim col As Byte, row As Byte, n As Byte
  Dim dayCounter As Byte
  Dim posX As Integer, posY As Integer
  Dim txtWidth As Integer, txtHeight As Integer
  
  With picBox
    txtWidth = .TextWidth("8") ' Text width
    txtHeight = .TextHeight("8") ' Text height
    
    ' Find out which day of the week the 1st of the month falls on
    col = Weekday("1/" + Str(currentMonth) + "/" + Str(currentYear), vbMonday) - 1
    row = 1 ' Start from row 1
    posY = picBox.CurrentY '/ Screen.TwipsPerPixelY
    
    ' ~~~~~~~~~~~~~ Write the weekdays with 2 letters ~~~~~~~~
    For n = 1 To 7
      PicPrint (n - 1) * txtWidth * 3 + x, posY, o, weekDays(n), weekDayColor
    Next

    '~~~~~~~~~~~~~~ Write the day numbers ~~~~~~~~~
    For dayCounter = 1 To 31
      If Not IsDate(dayCounter & "/" & currentMonth & "/" & currentYear) Then Exit For ' Exit if beyond the last day of the month
      ' If the row is full, move to the next row
      If col = 7 Then
        col = 0
        .CurrentY = row * txtHeight ' Move to the next row
        row = row + 1
      End If
      posX = x + col * txtWidth * 3
      posY = y + row * txtHeight
      .CurrentX = posX
      .CurrentY = posY
      col = col + 1
      
      ' If this is the current month's calendar
      If o = 2 Then
        If dayCounter = Day(Now) Then ' If today is the current date
          ' Draw a filled ellipse around the date
          Select Case currentDayShape
            Case "Circle"
              picBox.Circle (posX + txtWidth - 1, posY + txtHeight * 0.5), txtHeight * 0.65, fontColor
            Case "Ellipse"
              picBox.Circle (posX + txtWidth - 1, posY - 1 + txtHeight * 0.5), txtHeight * 0.5 + (txtHeight / 3), fontColor, , , 0.75
            Case "Rectangle"
              picBox.Line Step(-txtWidth / 2, 0)-Step(txtWidth * 2.75, txtHeight), fontColor, B
            Case "RoundRectangle"
              picBox.ForeColor = fontColor
              RoundRect picBox.hdc, posX - txtWidth * 0.5, posY, posX + txtWidth * 2.5, posY + txtHeight, txtHeight * 0.6, txtHeight * 0.6 'X1, Y1, X2, Y2, X radius, Y radius
          End Select
          .CurrentX = posX
          .CurrentY = posY
          picBox.Refresh
        End If
      End If
      
      ' If the day number is less than 10, adjust the position by one character
      If dayCounter < 10 Then
        posX = posX + txtWidth / 2
        .CurrentX = posX
      End If
      
      ' Print the day number
      PicPrint posX, posY, o, CStr(dayCounter), GetDayColor(dayCounter, currentMonth, currentYear, holidays)
    Next
  End With
End Sub

Private Function GetDayColor(dayCounter As Byte, currentMonth As Byte, currentYear As Integer, holidays As String) As Long
  ' Weekends are red
  If Weekday(dayCounter & "/" & currentMonth & "/" & currentYear, vbMonday) > 5 Then
    GetDayColor = holidayColor
  Else
    GetDayColor = fontColor
    ' If there are holidays
    If holidays <> "" And InStr(1, holidays, "," & CStr(dayCounter) & "/" & CStr(currentMonth) & ",") > 0 Then
      GetDayColor = holidayColor
    End If
  End If
End Function

Private Sub PicPrint(x As Integer, y As Integer, shadowOffset As Byte, text As String, txtColor As Long)
  With picBox
    .ForeColor = shadowColor
    If textEffect = "shadow" Or textEffect = "outline" Then
      .CurrentX = x + shadowOffset
      .CurrentY = y + shadowOffset
      picBox.Print text
    End If
    If textEffect = "outline" Then
      .CurrentX = x - shadowOffset
      .CurrentY = y - shadowOffset
      picBox.Print text
      .CurrentX = x + shadowOffset
      .CurrentY = y - shadowOffset
      picBox.Print text
      .CurrentX = x - shadowOffset
      .CurrentY = y + shadowOffset
      picBox.Print text
    End If
    .ForeColor = txtColor
    .CurrentX = x
    .CurrentY = y
    picBox.Print text
End With
End Sub

Private Function LoadHolidays() As String
  Dim holidaysFile As String
  holidaysFile = App.Path & "\" & holidaysFname
  
  ' Handle errors when opening the file
  On Error GoTo ErrorHandler
  If FileExists(holidaysFile) Then  ' If the holidays file exists
    Dim lineText As String, txtHoliday As String
    Open holidaysFile For Input As #1
    Do Until EOF(1)
      Line Input #1, lineText
      lineText = Trim(lineText)
    If IsDate(lineText & "/2024") Then txtHoliday = txtHoliday & lineText & ","
    Loop
    Close
    If txtHoliday <> "" Then LoadHolidays = "," & txtHoliday
  End If
  
  ' Disable error handling
  On Error GoTo 0
Exit Function

ErrorHandler:
  MsgBox holidaysFile & vbCr & "file could not be opened." & vbCr & "Error: " & Err.Description, vbCritical
  If Err.Number <> 0 Then Close #1
  On Error GoTo 0
End Function

Private Sub ApplyWallpaper(wallpaperFile As String)
  ' Desktop background color
    picBox.BackColor = GetSysColor(COLOR_DESKTOP)
    
  ' If a wallpaper is set
  If wallpaperFile <> "" Then
    Dim picBoxWidth As Integer, picBoxHeight As Integer
    picBoxWidth = picBox.ScaleWidth
    picBoxHeight = picBox.ScaleHeight
    
    On Error Resume Next
    img.Picture = LoadPicture(wallpaperFile)
    If Err Then
      Err.Clear
      Set img.Picture = LoadPicture_2(wallpaperFile)
      If Err Then
        MsgBox "Unable to load the wallpaper file." & vbCr & _
          "Please check the file path or format and try again." & vbCr & _
          Err.Number & " : " & Err.Description, , "Error"
        End
      End If
    End If
    On Error GoTo 0
    
    ' If the image dimensions are equal to the PictureBox dimensions
    If img.Width = picBoxWidth And img.Height = picBoxHeight Then
      picBox.Picture = img.Picture
    Else
      ' If the wallpaper is tiled
      If reg.ReadKey(regCpanelDeskopPath & "TileWallpaper") = "1" Then
        ' Tile the wallpaper
        Dim h1 As Integer, w1 As Integer
        For h1 = 0 To picBoxHeight Step img.Height
          For w1 = 0 To picBoxWidth Step img.Width
            picBox.PaintPicture img.Picture, w1, h1
          Next
        Next
      Else
        ' Determine wallpaper style
        Dim imgRatio As Double, picBoxRatio As Double
        Dim newHeight As Double, newWidth As Double
        Dim wpstil As Integer
        
        ' Wallpaper Style
        wpstil = Val(reg.ReadKey(regCpanelDeskopPath & "WallpaperStyle"))
        imgRatio = img.Width / img.Height
        picBoxRatio = picBoxWidth / picBoxHeight

        Select Case wpstil
          Case 0 ' Centered
            picBox.PaintPicture img.Picture, (picBoxWidth - img.Width) / 2, (picBoxHeight - img.Height) / 2
          Case 2  ' Stretched
            picBox.PaintPicture img.Picture, 0, 0, picBoxWidth, picBoxHeight, 0, 0, img.Width, img.Height
          Case 3, 6 ' Fit
            If imgRatio > picBoxRatio Then
                ' Width fits, height is adjusted to maintain aspect ratio
                newHeight = picBoxWidth / imgRatio
                picBox.PaintPicture img.Picture, 0, (picBoxHeight - newHeight) / 2, picBoxWidth, newHeight, 0, 0, img.Width, img.Height
            Else
                ' Height fits, width is adjusted to maintain aspect ratio
                newWidth = picBoxHeight * imgRatio
                picBox.PaintPicture img.Picture, (picBoxWidth - newWidth) / 2, 0, newWidth, picBoxHeight, 0, 0, img.Width, img.Height
            End If
          Case 10, 22 ' Fill
            If imgRatio < picBoxRatio Then
                ' Width fits, height overflows (Aspect ratio is larger for the image)
                newHeight = picBoxWidth / imgRatio
                picBox.PaintPicture img.Picture, 0, (picBoxHeight - newHeight) / 2, picBoxWidth, newHeight, 0, 0, img.Width, img.Height
            Else
                ' Height fits, width overflows (Aspect ratio is larger for the screen)
                newWidth = picBoxHeight * imgRatio
                picBox.PaintPicture img.Picture, (picBoxWidth - newWidth) / 2, 0, newWidth, picBoxHeight, 0, 0, img.Width, img.Height
            End If
          Case Else
            picBox.Picture = LoadPicture(wallpaperFile)
        End Select
      End If
    End If
  End If
End Sub

'Loads other image format files. PNG, TIFF
Private Function LoadPicture_2(strPath As String) As StdPicture
  With CreateObject("WIA.ImageFile")
    .LoadFile (strPath)
    Set LoadPicture_2 = .FileData.Picture
  End With
End Function

Private Function FileExists(ByVal sFileName As String) As Boolean
  Dim intReturn As Integer
  On Error GoTo FileExists_Error
  intReturn = GetAttr(sFileName)
  FileExists = True
  Exit Function

FileExists_Error:
  FileExists = False
End Function

