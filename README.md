# AS Desktop Calendar

This application generates a desktop background with calendars for the current and next month, displayed in the system’s locale language.

![Screenshot](https://github.com/akcansoft/Desktop-Calendar/blob/main/ss-1.jpg)

## Overview

The **AS Desktop Calendar** is a Visual Basic 6 (VB6) application that dynamically overlays a calendar onto the desktop wallpaper. It includes functionality for showing holidays and weekends and offers customization options for appearance.

## What's New in v1.1

- **Desktop Calendar Position Adjustment:** You can now customize the calendar's position on the desktop using `startOffsetX` and `startOffsetY` options in the INI file.
- **Improved Wallpaper Integration:** The application now fully aligns with Windows desktop placement options, ensuring seamless integration of the calendar with the wallpaper, whether tiled, centered, stretched, or using the fit or fill options.
- **Dynamic Image Scaling:** The application now better handles image aspect ratios, dynamically scaling the calendar based on the wallpaper's dimensions.
- **Active Day Marking Options:** New shape options have been added to mark the current day on the calendar. You can now choose between:
  - Circle
  - Ellipse
  - Rectangle
  - RoundRectangle
 
  These options can be customized through the INI file for a personalized look.
- **Bug Fixes and Performance Improvements:** Several optimizations to improve memory usage and stability during wallpaper generation.
- **PNG and TIF File Support:** The application can now generate calendars using PNG and TIF files as desktop wallpapers in addition to BMP and JPG.

## Key Features

- **Dynamic Wallpaper with Calendar:** Automatically generates a wallpaper featuring the current and next month’s calendars, displayed on the user's desktop wallpaper.
- **Holiday Highlighting:** Highlights holidays in red. The holidays are read from an external text file (`holidays.txt`).
- **Weekend Highlighting:** Automatically highlights weekends.
- **Existing Wallpaper Integration:** The generated calendar is integrated with the existing desktop wallpaper without overwriting the original image.
- **Customizable Settings:** Options for adjusting calendar position, font style, colors, and other appearance settings are defined in the INI file.
- **Locale Support:** Uses the system locale to display the calendar’s months and weekdays in the local language.

## Usage

### 1. Installation
- Copy the compiled executable file, `holidays.txt`, and  `setting.ini` file to the desired directory.

### 2. Configuring Holidays
- Add your holidays to the `holidays.txt` file in the `dd/mm` format, with one date per line (e.g., `25/12` for December 25th).

### 3. Configuring Settings

The `settings.ini` file allows you to customize various aspects of the calendar displayed on your desktop wallpaper. Below are the configuration options available:

#### [FONT]
- **fontName:** The name of the font used for the calendar text. Default is `Tahoma`.
- **fontBold:** Set to `True` to enable bold text, or `False` for regular text.
- **fontItalic:** Set to `True` to enable italic text, or `False` for normal text.
- **fontColor:** The color of the calendar text in hexadecimal format (e.g., `&HFFFFFF` for white).
- **shadowColor:** The color of the shadow effect on the text in hexadecimal format (e.g., `&H000000` for black).
- **weekdayColor:** The color used for weekdays text in hexadecimal format.
- **holidayColor:** The color used for holidays text in hexadecimal format.
- **fontRatio_1:** The ratio of the current month’s font height to the screen height. Default is `45`.
- **fontRatio_2:** The ratio of the next month’s font height to the screen height. Default is `65`.

#### [SHAPE]
- **currentDayShape:** Determines the shape used to highlight the current day. Options include `Circle`, `Ellipse`, `Rectangle`, and `RoundRectangle`.
- **shapeFillColor:** The fill color of the shape used for the current day, specified in hexadecimal format (e.g., `&H30B4F3`).

#### [CALENDAR POSITION]
- **startOffsetX:** The horizontal offset from the top center of the screen. Use positive or negative values to adjust the calendar’s position.
- **startOffsetY:** The vertical offset from the top center of the screen. Adjust the position by using positive or negative values.

These settings allow you to tailor the appearance and positioning of the calendar to match your preferences and desktop setup.


### 4. Running the Application
- Launch the executable to generate the wallpaper with the embedded calendar. The application automatically applies the generated wallpaper as the desktop background.
- To keep the calendar updated at each startup, place a shortcut of the executable in the Windows Startup folder.

## Files

- **`aswallpaper.bmp`:** The generated wallpaper file.
- **`holidays.txt`:** A text file containing a list of holidays.
- **`settings.ini`:** Contains customizable settings like font size, colors, and calendar position on the desktop.

## Dependencies

- Windows OS
- Visual Basic 6 Runtime

## Version History

- **First Release**: 16/04/2004
- **v1.0:**: 30/08/2024
- **v1.1:**: 06/09/2024
  
## Contribution

Contributions are welcome! You can submit a pull request for improvements or new features.

## License

This project is licensed under the GPL-3.0 License.

## Author

- **Mesut AKCAN**
  - Blog: [akcansoft.blogspot.com](http://akcansoft.blogspot.com)
  - YouTube: [youtube.com/mesutakcan](http://youtube.com/mesutakcan)
  - Email: [makcan@gmail.com](mailto:makcan@gmail.com)
