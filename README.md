# ConsoleUtilities
This is a module to set properties of the console that you normally set in the console's property dialog.

For example, if you want to set the RGB values for the 16 colors and set the text foreground and background colors, you could use:

```powershell
    Set-ConsoleColorTable -ColorTable @{
        Black       = '#141414'
        Blue        = '#004bff'
        Cyan        = '#00ffff'
        DarkBlue    = '#374b80'
        DarkCyan    = '#008080'
        DarkGray    = '#808080'
        DarkGreen   = '#008000'
        DarkMagenta = '#800080'
        DarkRed     = '#800000'
        DarkYellow  = '#808000'
        Gray        = '#c0c0c0'
        Green       = '#00ff00'
        Magenta     = '#ff00ff'
        Red         = '#ff0000'
        White       = '#ffffff'
        Yellow      = '#ffff00'
    } -TextForegroundColor Gray -TextBackgroundColor Black
```
