VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF972F&
      FillColor       =   &H0030B4F3&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Mesut AKCAN
' 16/04/2004
'
' Update: 30/08/2024
'
' http://akcansoft.blogspot.com
' http://youtube.com/mesutakcan
' makcan@gmail.com
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
  (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, _
  ByVal fuWinIni As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Const COLOR_DESKTOP As Long = 1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = 1
Const HKEY_CURRENT_USER = &H80000001
Const wallpaperFname As String = "\aswallpaper.bmp"
Const holidaysFname As String = "\holidays.txt"
'Font
Const txtFontName As String = "Tahoma"
Const txtFontBold As Boolean = True
Const fontRatio As Double = 0.00125
'Font colors
Const holidayColor As Long = vbRed
Const fontColor As Long = vbWhite
Const fontShadowColor As Long = vbBlack
Const weekdayFontColor As Long = &H15AE4F
Const circleFillColor As Long = &H30B4F3

Dim currentMonth As Byte, currentYear As Integer
Dim weeks(1 To 7) As String * 2
Dim holidays As String

Private Sub Form_Load()
  Dim n As Byte
  Dim orjDesktopWallpaperFile As String
  Dim months(1 To 12) As String, dateText As String
  Dim dateTxtWidth As Integer, offsetX As Integer, offsetY As Integer
  Dim holidaysFile As String
  Dim posX As Integer, posY As Integer

  orjDesktopWallpaperFile = QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")

  ' Desktop background color
  picBox.BackColor = GetSysColor(COLOR_DESKTOP)

  ' If a desktop wallpaper is set
  If orjDesktopWallpaperFile <> "" Then
    ' If the desktop wallpaper is the same as the one in the application directory
    If orjDesktopWallpaperFile = App.Path & wallpaperFname Then
      ' Retrieve the previously saved desktop wallpaper from the registry
      orjDesktopWallpaperFile = GetSetting(App.EXEName, "PrevWallPaper", "File")
    
    ' If the desktop wallpaper has changed
    Else
      ' Save the original desktop wallpaper file to the registry
      SaveSetting App.EXEName, "PrevWallPaper", "File", orjDesktopWallpaperFile
    End If
  End If

' Create a list of month names in the local language
For n = 1 To 12
  months(n) = Format("1/" & n & "/2024", "MMMM") ' Year is irrelevant
Next

' Create a list of weekdays in the local language. The first date provided is Monday
For n = 1 To 7
  weeks(n) = Format(n & "/7/2024", "ddd") ' 01/07/2024 -> Monday
Next

' Load holidays from the file
holidays = LoadHolidays()

Me.Width = Screen.Width
Me.Height = Screen.Height
Me.Move 0, 0

currentMonth = Month(Now()) ' Current month
currentYear = Year(Now()) ' Current year

With picBox
  .Width = Screen.Width / Screen.TwipsPerPixelX
  .Height = Screen.Height / Screen.TwipsPerPixelY
  .Move 0, 0
  .Font.Name = txtFontName
  .Font.Size = Int(Screen.Height * fontRatio)
  .FontBold = txtFontBold
  .FillColor = circleFillColor
  
  ApplyWallpaper orjDesktopWallpaperFile
  
  '~~~~~~~~~~~~ Write today's date ~~~~~~~~~~~~
  dateText = Format(Now, "d MMMM yyyy dddd")
  dateTxtWidth = .TextWidth(dateText)
  offsetX = (.Width / 2) - (dateTxtWidth / 2) ' Center the text on the screen
  offsetY = 10
  PicPrint offsetX, offsetY, 2, dateText, fontColor
  
  '~~~~~~~~~~~~ This month ~~~~~~~~~~~~~~~~~~~~
  Calendar offsetX, .CurrentY, 2
  
  '~~~~~~~~~~~~ Next month ~~~~~~~~~~~~~~~~~~~~
  .Font.Size = Int(.Font.Size / 2) ' Reduce font size for next month's text
  .CurrentX = offsetX
  .CurrentY = .CurrentY + .TextHeight("8") / 2 ' Add half the text height
  
  ' If the current month is December, the next month will be January of the following year
  currentMonth = currentMonth + 1
  
  If currentMonth > 12 Then
    currentMonth = 1
    currentYear = currentYear + 1
  End If
  
  posX = .CurrentX
  posY = .CurrentY
  PicPrint posX, posY, 1, months(currentMonth) & " " & currentYear, fontColor ' Next month's header
  
  '~~~~~~~~~~~~ Next month calendar ~~~~~~~~~~~~
  Calendar offsetX, .CurrentY, 1
  
  '~~~~~ Write "akcanSoft" ~~~~~~~~~~
  .CurrentX = posX
  PicPrint posX, .CurrentY, 1, "akcanSoft", &HCCCCCC
End With

Dim wallpaperFileName As String
wallpaperFileName = App.Path & wallpaperFname
SavePicture picBox.Image, wallpaperFileName
' Set the saved image file as the desktop wallpaper
SystemParametersInfo SPI_SETDESKWALLPAPER, 0, wallpaperFileName, SPIF_UPDATEINIFILE
End
End Sub

'~~~~~~~~~~~~ CALENDAR ~~~~~~~~~~~~~~~~~~
Sub Calendar(offsetX, offsetY, o As Byte)
  Dim col As Byte, row As Byte, n As Byte
  Dim dayCounter As Byte, posX As Integer, posY As Integer
  Dim txtWidth As Integer, txtHeight As Integer
  
  With picBox
    txtWidth = .TextWidth("8") ' Text width
    txtHeight = .TextHeight("8") ' Text height
    
    ' Find out which day of the week the 1st of the month falls on
    col = Weekday("1/" + Str(currentMonth) + "/" + Str(currentYear), vbMonday) - 1
    row = 1 ' Start from row 1
    posY = .CurrentY
    
    ' ~~~~~~~~~~~~~ Write the weekdays with 2 letters ~~~~~~~~
    For n = 1 To 7
      PicPrint (n - 1) * txtWidth * 3 + offsetX, posY, o, weeks(n), weekdayFontColor
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
      posX = col * txtWidth * 3 + offsetX
      posY = row * txtHeight + offsetY
      .CurrentX = posX
      .CurrentY = posY
      col = col + 1
      
      ' If this is the current month's calendar
      If o = 2 Then
        If dayCounter = Day(Now) Then ' If today is the current date
          ' Draw a filled ellipse around the date
          picBox.Circle (posX + txtWidth - 1, posY - 1 + txtHeight / 2), txtHeight / 2 + (txtHeight / 3), fontColor, , , 0.75
          .CurrentX = posX
          .CurrentY = posY
        End If
      End If
      
      ' If the day number is less than 10, adjust the position by one character
      If dayCounter < 10 Then
        posX = posX + txtWidth
        .CurrentX = posX
      End If
      
      ' Print the day number
      PicPrint posX, posY, o, CStr(dayCounter), GetDayColor(dayCounter, currentMonth, currentYear, holidays)
    Next
  End With
End Sub

Function GetDayColor(dayCounter As Byte, currentMonth As Byte, currentYear As Integer, holidays As String) As Long
  ' Weekends are red
  If Weekday(dayCounter & "/" & currentMonth & "/" & currentYear, vbMonday) > 5 Then
    GetDayColor = holidayColor
  Else
    GetDayColor = fontColor
    ' If there are holidays
    If holidays <> "" And InStr(1, holidays, CStr(dayCounter) & "/" & CStr(currentMonth)) > 0 Then
      GetDayColor = holidayColor
    End If
  End If
End Function

Sub PicPrint(x As Integer, y As Integer, shadowOffset As Byte, text As String, txtColor As Long)
  With picBox
    .ForeColor = fontShadowColor
    .CurrentX = x
    .CurrentY = y
    picBox.Print text
    .ForeColor = txtColor
    .CurrentX = x - shadowOffset
    .CurrentY = y - shadowOffset
  picBox.Print text
End With
End Sub

Private Function LoadHolidays() As String
  Dim holidaysFile As String
  holidaysFile = App.Path & holidaysFname
  
  ' Handle errors when opening the file
  On Error GoTo ErrorHandler
  If Dir(holidaysFile) <> "" Then ' If the holidays file exists
    Dim lineText As String, txtHoliday As String
    Open holidaysFile For Input As #1
    Do Until EOF(1)
      Line Input #1, lineText
      lineText = Trim(lineText)
    If IsDate(lineText & "/2024") Then txtHoliday = txtHoliday & lineText & ","
    Loop
    Close
  LoadHolidays = txtHoliday
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
  ' If a wallpaper is set
  If wallpaperFile <> "" Then
    Dim img As Image
    Dim h1, w1
    Dim pWidth, pHeight
    pWidth = picBox.ScaleWidth
    pHeight = picBox.ScaleHeight
    
    Set img = Me.Controls.Add("VB.Image", "imgTemp")
    img.Picture = LoadPicture(wallpaperFile)
    
    ' If the wallpaper is tiled
    If QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper") = "1" Then
      For h1 = 0 To pHeight Step img.Height
        For w1 = 0 To pWidth Step img.Width
          picBox.PaintPicture img.Picture, w1, h1
        Next
      Next
    
    ' If the wallpaper is centered or stretched
    Else
      Dim wpstil As String
      ' WallpaperStyle : Value Data
      ' Center 0 'Fill 4 'Fit 3 'Span 5 'Stretch 2  'Tile 1
      wpstil = QueryValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle")
      If wpstil = "0" Then ' Centered
        picBox.PaintPicture img.Picture, (pWidth - img.Width) / 2, (pHeight - img.Height) / 2
      ElseIf wpstil = "2" Then ' Stretched
        picBox.PaintPicture img.Picture, 0, 0, pWidth, pHeight, 0, 0, img.Width, img.Height
      Else
        picBox.Picture = LoadPicture(wallpaperFile)
      End If
    End If
    
    Me.Controls.Remove "imgTemp"
    Set img = Nothing
  End If
End Sub


