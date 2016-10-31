using System;
using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Runtime.InteropServices;

namespace ConsoleUtilities
{
    [StructLayout(LayoutKind.Sequential)]
    public struct COORD
    {
        public short X;
        public short Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SMALL_RECT
    {
        public short Left;
        public short Top;
        public short Right;
        public short Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RGBColor
    {
        public RGBColor(int red, int green, int blue)
        {
            if (red < 0 || red > 255 || green < 0 || green > 255 || blue < 0 || blue > 255)
            {
                 throw new ArgumentException();
            }

            RGB = (red | (green << 8) | blue << 16);
        }

        public RGBColor(int rgb)
        {
            if (rgb < 0 || rgb > 0xffffff)
                throw new ArgumentException();

            RGB = SwapRedAndBlue(rgb);
        }

        public RGBColor(string rgb)
        {
            if (string.IsNullOrWhiteSpace(rgb))
                throw new ArgumentException();

            if (rgb[0] == '#')
                rgb = rgb.Substring(1);

            int val;
            int.TryParse(rgb, NumberStyles.HexNumber, NumberFormatInfo.InvariantInfo, out val);

            if (val < 0 || val > 0xffffff)
                throw new ArgumentException();

            RGB = SwapRedAndBlue(val);
        }

        private static int SwapRedAndBlue(int val)
        {
            return ((val & 0xff) << 16) | (val & 0xff00) | ((val & 0xff0000) >> 16);
        }

        public int RGB;

        public int R => RGB & 0xff;

        public int G => (RGB >> 8) & 0xff;

        public int B => (RGB >> 16) & 0xff;

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "#{0:x2}{1:x2}{2:x2}", R, G, B);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CONSOLE_SCREEN_BUFFER_INFOEX
    {
        public int cbSize;
        public COORD dwSize;
        public COORD dwCursorPosition;
        public short wAttributes;
        public SMALL_RECT srWindow;
        public COORD dwMaximumWindowSize;
        public short wPopupAttributes;
        [MarshalAs(UnmanagedType.Bool)]
        public bool bFullscreenSupported;
        public RGBColor Black;
        public RGBColor DarkBlue;
        public RGBColor DarkGreen;
        public RGBColor DarkCyan;
        public RGBColor DarkRed;
        public RGBColor DarkMagenta;
        public RGBColor DarkYellow;
        public RGBColor Gray;
        public RGBColor DarkGray;
        public RGBColor Blue;
        public RGBColor Green;
        public RGBColor Cyan;
        public RGBColor Red;
        public RGBColor Magenta;
        public RGBColor Yellow;
        public RGBColor White;
    }

    public class NativeMethods
    {
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetConsoleScreenBufferInfoEx(
            IntPtr hConsoleOutput,
            ref CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetConsoleScreenBufferInfoEx(
            IntPtr hConsoleOutput,
            ref CONSOLE_SCREEN_BUFFER_INFOEX lpConsoleScreenBufferInfoEx);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr GetStdHandle(int handleId);
    }

    [Cmdlet("Set", "ConsoleColorTable")]
    public class SetConsoleColorTableCommand : PSCmdlet
    {
        [Parameter(ParameterSetName = "SingleColor")]
        public RGBColor RGB { get; set; }

        [Parameter(ParameterSetName = "SingleColor")]
        public ConsoleColor ConsoleColor { get; set; }

        [Parameter(ParameterSetName = "SetMultipleColors")]
        public Hashtable ColorTable { get; set; }

        [Parameter]
        public ConsoleColor? TextForegroundColor { get; set; }
        [Parameter]
        public ConsoleColor? TextBackgroundColor { get; set; }

        [Parameter]
        public ConsoleColor? PopupForegroundColor { get; set; }

        [Parameter]
        public ConsoleColor? PopupBackgroundColor { get; set; }

        protected override void EndProcessing()
        {
            var stdOutHandle = NativeMethods.GetStdHandle(-11);

            CONSOLE_SCREEN_BUFFER_INFOEX screenBufferInfo = new CONSOLE_SCREEN_BUFFER_INFOEX 
            {
                cbSize = Marshal.SizeOf<CONSOLE_SCREEN_BUFFER_INFOEX>()
            };
            if (!NativeMethods.GetConsoleScreenBufferInfoEx(stdOutHandle, ref screenBufferInfo))
            {
                WriteError(new ErrorRecord(
                    null,
                    "ErrorGettingScreenBufferInfo",
                    ErrorCategory.InvalidOperation,
                    "Error getting screen buffer info"));
                return;
            }

            if (ParameterSetName == "SingleColor")
            {
                SetColorHelper(ConsoleColor, RGB, ref screenBufferInfo);
            }
            else
            {
                var pairs = ColorTable.GetEnumerator();
                while (pairs.MoveNext())
                {
                    var consoleColor = pairs.Key as ConsoleColor? ?? LanguagePrimitives.ConvertTo<ConsoleColor>(pairs.Key);
                    var color = pairs.Value as RGBColor? ?? LanguagePrimitives.ConvertTo<RGBColor>(pairs.Value);
                    SetColorHelper(consoleColor, color, ref screenBufferInfo);
                }
            }

            if (TextForegroundColor.HasValue)
            {
                screenBufferInfo.wAttributes = (short)((screenBufferInfo.wAttributes & 0xfff0) | (int)TextForegroundColor.Value);
            }

            if (TextBackgroundColor.HasValue)
            {
                screenBufferInfo.wAttributes = (short)((screenBufferInfo.wAttributes & 0xff0f) | (int)TextBackgroundColor.Value << 4);
            }

            if (PopupForegroundColor.HasValue)
            {
                screenBufferInfo.wPopupAttributes = (short)((screenBufferInfo.wPopupAttributes & 0xfff0) | (int)PopupForegroundColor.Value);
            }

            if (PopupBackgroundColor.HasValue)
            {
                screenBufferInfo.wPopupAttributes = (short)((screenBufferInfo.wPopupAttributes & 0xff0f) | (int)PopupBackgroundColor.Value << 4);
            }

            // conhost returns an off-by-one window size, so we must compensate
            screenBufferInfo.srWindow.Bottom += 1;
            screenBufferInfo.srWindow.Right += 1;
            NativeMethods.SetConsoleScreenBufferInfoEx(stdOutHandle, ref screenBufferInfo);
        }

        private void SetColorHelper(ConsoleColor consoleColor, RGBColor color, ref CONSOLE_SCREEN_BUFFER_INFOEX screenBufferInfo)
        {
            switch (consoleColor)
            {
                case ConsoleColor.Black:       screenBufferInfo.Black = color; break;
                case ConsoleColor.DarkBlue:    screenBufferInfo.DarkBlue = color; break;
                case ConsoleColor.DarkGreen:   screenBufferInfo.DarkGreen = color; break;
                case ConsoleColor.DarkCyan:    screenBufferInfo.DarkCyan = color; break;
                case ConsoleColor.DarkRed:     screenBufferInfo.DarkRed = color; break;
                case ConsoleColor.DarkMagenta: screenBufferInfo.DarkMagenta = color; break;
                case ConsoleColor.DarkYellow:  screenBufferInfo.DarkYellow = color; break;
                case ConsoleColor.Gray:        screenBufferInfo.Gray = color; break;
                case ConsoleColor.DarkGray:    screenBufferInfo.DarkGray = color; break;
                case ConsoleColor.Blue:        screenBufferInfo.Blue = color; break;
                case ConsoleColor.Green:       screenBufferInfo.Green = color; break;
                case ConsoleColor.Cyan:        screenBufferInfo.Cyan = color; break;
                case ConsoleColor.Red:         screenBufferInfo.Red = color; break;
                case ConsoleColor.Magenta:     screenBufferInfo.Magenta = color; break;
                case ConsoleColor.Yellow:      screenBufferInfo.Yellow = color; break;
                case ConsoleColor.White:       screenBufferInfo.White = color; break;
            }
        }
    }

    [Cmdlet("Get", "ConsoleColorTable")]
    public class GetConsoleColorTableCommand : PSCmdlet
    {
        protected override void EndProcessing()
        {
            var stdOutHandle = NativeMethods.GetStdHandle(-11);

            CONSOLE_SCREEN_BUFFER_INFOEX screenBufferInfo = new CONSOLE_SCREEN_BUFFER_INFOEX 
            {
                cbSize = Marshal.SizeOf<CONSOLE_SCREEN_BUFFER_INFOEX>()
            };
            NativeMethods.GetConsoleScreenBufferInfoEx(stdOutHandle, ref screenBufferInfo);

            WriteObject(screenBufferInfo);
        }
    }
}
