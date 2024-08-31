# AS Desktop Calendar
This application makes the current month and next month calendar in local language on the desktop image.
![Screenshot](https://github.com/akcansoft/Desktop-Calendar/blob/main/ss-1.jpg)

This project is a Visual Basic 6 (VB6) application that dynamically generates a desktop wallpaper with an embedded calendar.
It displays the current date, a calendar for the current and next month, and overlays this on the user's existing wallpaper.
The application also highlights holidays and weekends.

## Features

- **Dynamic Wallpaper Generation:** Generates a new wallpaper based on the current date and the calendar for the current and next month.
- **Holiday Highlighting:** Highlights holidays in red. Holidays are loaded from an external text file (`holidays.txt`).
- **Existing Wallpaper Integration:** Integrates the generated calendar with the user's existing desktop wallpaper, preserving the original wallpaper's style (e.g., tiled, centered, or stretched).
- **Customizable Appearance:** Font, colors, and positioning are customizable through constants defined in the code.
- **Supports Locale:** Displays months and weekdays in the system's locale language.

## Usage

1. **Installation:**
   - Copy the compiled executable and the `holidays.txt` file to the desired directory.

2. **Holiday Configuration:**
   - The `holidays.txt` file should be placed in the same directory as the executable.
   - Each holiday should be listed in the `dd/mm` format, one per line (e.g., `25/12` for December 25th, `5/1` for January 5th).

3. **Running the Application:**
   - Run the compiled executable. The application will automatically generate a new wallpaper and set it as the desktop background.
   - The application does not run continuously in memory. Therefore, place the application shortcut in the Start/All programs/Startup folder so that the calendar is updated every time windows starts.

4. **Reverting to Original Wallpaper:**
   - The application saves the original wallpaper before applying the new one. To revert, simply close the application, and the original wallpaper will be restored.

## Files

- **`aswallpaper.bmp`:** The generated wallpaper file with the calendar overlay.
- **`holidays.txt`:** A text file containing the list of holidays.

## Dependencies

- Windows operating system.
- Visual Basic 6 runtime.

## Contribution
Contributions are always welcome! You can create a **pull request** to fix bugs or add new features.

## License

This project is open-source under the GPL-3.0 License.

## Author

- **Mesut AKCAN**  
  - Blog: [akcansoft.blogspot.com](http://akcansoft.blogspot.com)  
  - YouTube: [youtube.com/mesutakcan](http://youtube.com/mesutakcan)  
  - Email: [makcan@gmail.com](mailto:makcan@gmail.com)

## Updates

- **First Release:** 16/04/2004
- **Latest Update:** 30/08/2024

