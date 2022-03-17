using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.IO;
using Windows.Media.Ocr;
using System.Threading.Tasks;
using Windows.Storage;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace SimplyFlowWindows
{
    public class WindowsAutomator
    {
        // Internal timer
        int timer = 1000;
        // Time in ms between keystrokes
        int keyDelay = 100;
        // Time in ms between mouse events
        int mouseDelay = 100;
        // Audit message writer
        private StreamWriter auditWriter;
        // Log message writer
        private StreamWriter logWriter;
        // Process to be automated
        private Process automateProcess;
        // Working directory in which screen captures, logs and audit events are stored
        private string workingDirectory;
        // Filename of audit file
        private string auditFileName;
        // Filename of log file
        private string logFileName;
        // Process name
        private string processName = "empty";
        // Event list to be automated
        private List<AutomationEvent> events = new List<AutomationEvent>();

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        // Below links to native Windows functionality
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool GetCursorPos(out MousePoint lpMousePoint);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetWindowRect(IntPtr hWnd, ref RECT rect);
        public const int SRCCOPY = 0x00CC0020;

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest,
            int nWidth, int nHeight, IntPtr hObjectSource,
            int nXSrc, int nYSrc, int dwRop);
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth,
            int nHeight);
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteDC(IntPtr hDC);
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);


        // Modifier key definitions
        public static byte AltKey = 0x12;
        public static byte CtrlKey = 0x11;
        public static byte ShiftKey = 0x10;

        // Standard key definitions
        public static byte QKey = 0x10;
        public static byte WKey = 0x11;
        public static byte EKey = 0x12;
        public static byte RKey = 0x13;
        public static byte TKey = 0x14;
        public static byte YKey = 0x15;
        public static byte UKey = 0x16;
        public static byte IKey = 0x17;
        public static byte OKey = 0x18;
        public static byte PKey = 0x19;

        public static byte AKey = 0x1e;
        public static byte SKey = 0x1f;
        public static byte DKey = 0x20;
        public static byte FKey = 0x21;
        public static byte GKey = 0x22;
        public static byte HKey = 0x23;
        public static byte JKey = 0x24;
        public static byte KKey = 0x25;
        public static byte LKey = 0x26;

        public static byte ZKey = 0x2c;
        public static byte XKey = 0x2d;
        public static byte CKey = 0x2e;
        public static byte VKey = 0x2f;
        public static byte BKey = 0x30;
        public static byte NKey = 0x31;
        public static byte MKey = 0x32;

        public static byte Number1Key = 0x02;
        public static byte Number2Key = 0x03;
        public static byte Number3Key = 0x04;
        public static byte Number4Key = 0x05;
        public static byte Number5Key = 0x06;
        public static byte Number6Key = 0x07;
        public static byte Number7Key = 0x08;
        public static byte Number8Key = 0x09;
        public static byte Number9Key = 0x0a;
        public static byte Number0Key = 0x0b;

        public static byte WindowsKey = 0x5B;
        public static byte UpKey = 0x26;
        public static byte DownKey = 0x28;
        public static byte LeftKey = 0x25;
        public static byte RightKey = 0x27;
        public static byte EnterKey = 0x0D;
        public static byte TabKey = 0x09;
        public static byte SpaceBarKey = 0x20;
        public static byte HomeKey = 0x24;
        public static byte PgUpKey = 0x21;
        public static byte PgDnKey = 0x22;
        public static byte EndKey = 0x23;
        public static byte DelKey = 0x2e;
        public static byte InsKey = 0x2d;
        public static byte PrintScreenKey = 0x2c;
        public static byte EscKey = 0x1b;
        public static byte BackspaceKey = 0x08;
        public static byte CapsKey = 0x14;
        public static byte F1Key = 0x70;
        public static byte F2Key = 0x71;
        public static byte F3Key = 0x72;
        public static byte F4Key = 0x73;
        public static byte F5Key = 0x74;
        public static byte F6Key = 0x75;
        public static byte F7Key = 0x76;
        public static byte F8Key = 0x77;
        public static byte F9Key = 0x78;
        public static byte F10Key = 0x79;
        public static byte F11Key = 0x7a;
        public static byte F12Key = 0x7b;
        public static byte TildaKey = 0xc0;
        public static byte MinusKey = 0xbd;
        public static byte EqualsKey = 0xbb;
        public static byte OpenSquareKey = 0xdb;
        public static byte CloseSquareKey = 0xdd;
        public static byte BackslashKey = 0xdc;
        public static byte SemicolonKey = 0xba;
        public static byte SingleQuoteKey = 0xde;
        public static byte CommaKey = 0xbc;
        public static byte FullstopKey = 0xbe;
        public static byte ForwardslashKey = 0xbf;

        // Key up modifier
        public static uint KEYEVENTF_KEYUP = 0x0002;

        // Mouse event flags
        public enum MouseEventFlags
        {
            LeftDown = 0x00000002,
            LeftUp = 0x00000004,
            MiddleDown = 0x00000020,
            MiddleUp = 0x00000040,
            Move = 0x00000001,
            Absolute = 0x00008000,
            RightDown = 0x00000008,
            RightUp = 0x00000010
        }

        // Is the event a mouse or keyboard event
        public enum AutomationEventType
        {
            Keyboard = 0x00,
            Mouse = 0x01
        }

        /// <summary>
        /// Set cursor position of mouse
        /// </summary>
        /// <param name="point">Mousepoint x & y</param>
        private void SetCursorPosition(MousePoint point)
        {
            SetCursorPos(point.X, point.Y);
        }

        /// <summary>
        /// Get mouse cursor position
        /// </summary>
        /// <returns>Mousepoint</returns>
        public MousePoint GetCursorPosition()
        {
            MousePoint currentMousePoint;
            var gotPoint = GetCursorPos(out currentMousePoint);
            if (!gotPoint) { currentMousePoint = new MousePoint(0, 0); }
            return currentMousePoint;
        }

        /// <summary>
        /// Create mouse event
        /// </summary>
        /// <param name="value">Mouse event flags</param>
        private void MouseEvent(MouseEventFlags value)
        {
            MousePoint position = GetCursorPosition();

            mouse_event
                ((int)value,
                 position.X,
                 position.Y,
                 0,
                 0)
                ;
        }

        /// <summary>
        /// Audit a command - used mainly for logging and tracing
        /// </summary>
        /// <param name="commandName">Command name</param>
        /// <param name="args">Related arguments/info</param>
        private void AuditCommand(string commandName, string args)
        {
            auditWriter.WriteLine("Time:'" + timer + "' Command:'" + commandName + "' Arguments:'" + args + "'");
            auditWriter.Flush();
        }

        /// <summary>
        /// Log a message - used mainly for well formed user output
        /// </summary>
        /// <param name="moduleName">Module name</param>
        /// <param name="message">Message</param>
        public void LogMessage(string moduleName, string message)
        {
            logWriter.WriteLine("LogMessage module='" + moduleName + "' message='" + message + "'");
            logWriter.Flush();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="processName">Name of process being automated</param>
        /// <param name="workingDirectory">Working directory for audit/logging/screen capture</param>
        public WindowsAutomator(string processName, string workingDirectory)
        {
            this.workingDirectory = workingDirectory;
            this.processName = processName;
            auditFileName = workingDirectory + processName.Replace(" ", "") + DateTime.Now.ToShortDateString().Replace("/", "") + Guid.NewGuid().ToString() + ".txt";
            logFileName = workingDirectory + "userlog"+processName.Replace(" ", "") + DateTime.Now.ToShortDateString().Replace("/", "") + Guid.NewGuid().ToString() + ".txt";
            Init();
            AuditCommand("Constructor", processName);
        }

        /// <summary>
        /// Initialize automation
        /// </summary>
        private void Init()
        {
            try
            {
                if (auditWriter == null)
                {
                    auditWriter = new StreamWriter(auditFileName);
                }
                if (logWriter == null)
                {
                    logWriter = new StreamWriter(logFileName);
                }
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message + "\n Cannot initialize.");
                return;
            }

            AuditCommand("Init", "Working directory:'" + workingDirectory + "' audit file name:'" + auditFileName);
        }

        /// <summary>
        /// Is the process that was started running
        /// </summary>
        /// <returns>bool</returns>
        public bool IsProcessRunning()
        {
            bool isRunning = !automateProcess.HasExited;

            AuditCommand("IsProcessRunning", "" + isRunning);

            return (isRunning);
        }

        /// <summary>
        /// Focus on window of process being automated
        /// </summary>
        public void FocusProcess()
        {
            if ((automateProcess != null) && (!automateProcess.HasExited))
            {
                SetForegroundWindow(automateProcess.MainWindowHandle);
            }
        }

        /// <summary>
        /// Write keystroke
        /// </summary>
        /// <param name="key">Key as per definitions</param>
        private void WriteKeyStroke(byte key)
        {
            events.Add(new AutomationEvent(key, (byte)(key + 128), 0, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(key, (byte)(key + 128), KEYEVENTF_KEYUP, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
        }

        /// <summary>
        /// Write ctrl plus key
        /// </summary>
        /// <param name="key">String version of key</param>
        public void WriteCtrlPlusKey(string key)
        {
            WriteCtrlPlusKey((byte)key.ToUpper().ToCharArray()[0]);
            AuditCommand("WriteCtrlPlusKey", "" + key);
        }

        /// <summary>
        /// Write alt plus key
        /// </summary>
        /// <param name="key">String version of key</param>
        public void WriteAltPlusKey(string key)
        {
            WriteAltPlusKey((byte)key.ToUpper().ToCharArray()[0]);
            AuditCommand("WriteAltPlusKey", "" + key);
        }

        /// <summary>
        /// Write shift plus key
        /// </summary>
        /// <param name="key">String version of key</param>
        public void WriteShiftPlusKey(string key)
        {
            WriteShiftPlusKey((byte)key.ToCharArray()[0]);
            AuditCommand("WriteShiftPlusKey", "" + key);
        }

        /// <summary>
        /// Write shift plus key
        /// </summary>
        /// <param name="key">Key as per definitions</param>
        public void WriteShiftPlusKey(byte key)
        {
            events.Add(new AutomationEvent(ShiftKey, (byte)(ShiftKey + 128), 0, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(key, (byte)(key + 128), 0, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(key, (byte)(key + 128), KEYEVENTF_KEYUP, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(ShiftKey, (byte)(ShiftKey + 128), KEYEVENTF_KEYUP, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            AuditCommand("WriteShiftPlusKey", "" + key);
        }

        /// <summary>
        /// Write alt plus key
        /// </summary>
        /// <param name="key">Key as per defintions</param>
        public void WriteAltPlusKey(byte key)
        {
            events.Add(new AutomationEvent(AltKey, (byte)(AltKey + 128), 0, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(key, (byte)(key + 128), 0, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(key, (byte)(key + 128), KEYEVENTF_KEYUP, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(AltKey, (byte)(AltKey + 128), KEYEVENTF_KEYUP, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            AuditCommand("WriteAltPlusKey", "" + key);
        }

        /// <summary>
        /// Write ctrl plus key
        /// </summary>
        /// <param name="key">Key as per definitions</param>
        public void WriteCtrlPlusKey(byte key)
        {
            events.Add(new AutomationEvent(CtrlKey, (byte)(AltKey + 128), 0, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(key, (byte)(key + 128), 0, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(key, (byte)(key + 128), KEYEVENTF_KEYUP, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            events.Add(new AutomationEvent(CtrlKey, (byte)(AltKey + 128), KEYEVENTF_KEYUP, 0, keyDelay, AutomationEventType.Keyboard));
            timer += keyDelay;
            AuditCommand("WriteCtrlPlusKey", "" + key);
        }

        /// <summary>
        /// Start a process
        /// </summary>
        /// <param name="filename">Filename to start</param>
        /// <param name="args">Associated arguments</param>
        public void StartProcess(string filename, string args)
        {
            // Start process and focus on it
            automateProcess = Process.Start(filename, args);
            SetForegroundWindow(automateProcess.MainWindowHandle);
            AuditCommand("StartProcess", "Filename:'" + filename + "' Arguments:'" + args + "'");
        }

        /// <summary>
        /// Run all events in automation list
        /// 
        /// It's best practice to run short event lists, check for expected result and then continue
        /// </summary>
        public void Run()
        {
            AuditCommand("Run", "");

            if ((automateProcess != null) && (!automateProcess.HasExited))
            {
                SetForegroundWindow(automateProcess.MainWindowHandle);
            }

            foreach (AutomationEvent automationEvent in events)
            {
                Thread.Sleep(automationEvent.Delay);

                if (automationEvent.Type == AutomationEventType.Keyboard)
                {
                    if (automationEvent.Extra != -1)
                    {
                        keybd_event(automationEvent.Vk, automationEvent.Scan, automationEvent.KeyEventInfo, automationEvent.Extra);
                    }
                }
                if (automationEvent.Type == AutomationEventType.Mouse)
                {
                    if (automationEvent.mouseEventFlags == MouseEventFlags.Move)
                    {
                        SetCursorPosition(automationEvent.mousePoint);
                    } else
                    {
                        MouseEvent(automationEvent.mouseEventFlags);
                    }
                }
            }

            events.Clear();

        }

        /// <summary>
        /// Write a string to keyboard buffer
        /// </summary>
        /// <param name="stringToWrite">String to write</param>
        public void WriteString(string stringToWrite)
        {
            foreach (byte b in Encoding.ASCII.GetBytes(stringToWrite))
            {
                if (((b >= 'a') && (b <= 'z')) ||
                    ((b >= '0') && (b <= '9')) ||
                    (b == SpaceBarKey))
                {
                    byte[] byteArray = new byte[1];
                    byteArray[0] = b;
                    string upper = Encoding.ASCII.GetString(byteArray).ToUpper();
                    WriteKeyStroke(Encoding.ASCII.GetBytes(upper)[0]);
                }

                if (((b >= 'A') && (b <= 'Z')))
                {
                    WriteShiftPlusKey(b);
                }

                int intChar = (int)b;
                switch (intChar)
                {
                    case '`':
                        WriteKeyStroke(TildaKey);
                        break;
                    case '-':
                        WriteKeyStroke(MinusKey);
                        break;
                    case '=':
                        WriteKeyStroke(EqualsKey);
                        break;
                    case '[':
                        WriteKeyStroke(OpenSquareKey);
                        break;
                    case ']':
                        WriteKeyStroke(CloseSquareKey);
                        break;
                    case 0x5c:
                        WriteKeyStroke(BackslashKey);
                        break;
                    case ';':
                        WriteKeyStroke(SemicolonKey);
                        break;
                    case 0x27:
                        WriteKeyStroke(SingleQuoteKey);
                        break;
                    case ',':
                        WriteKeyStroke(CommaKey);
                        break;
                    case '.':
                        WriteKeyStroke(FullstopKey);
                        break;
                    case '/':
                        WriteKeyStroke(ForwardslashKey);
                        break;
                    case '~':
                        WriteShiftPlusKey(TildaKey);
                        break;
                    case '!':
                        WriteShiftPlusKey((byte)'1');
                        break;
                    case '@':
                        WriteShiftPlusKey((byte)'2');
                        break;
                    case '#':
                        WriteShiftPlusKey((byte)'3');
                        break;
                    case '$':
                        WriteShiftPlusKey((byte)'4');
                        break;
                    case '%':
                        WriteShiftPlusKey((byte)'5');
                        break;
                    case '^':
                        WriteShiftPlusKey((byte)'6');
                        break;
                    case '&':
                        WriteShiftPlusKey((byte)'7');
                        break;
                    case '*':
                        WriteShiftPlusKey((byte)'8');
                        break;
                    case '(':
                        WriteShiftPlusKey((byte)'9');
                        break;
                    case ')':
                        WriteShiftPlusKey((byte)'0');
                        break;
                    case '_':
                        WriteShiftPlusKey(MinusKey);
                        break;
                    case '+':
                        WriteShiftPlusKey(EqualsKey);
                        break;
                    case '{':
                        WriteShiftPlusKey(OpenSquareKey);
                        break;
                    case '}':
                        WriteShiftPlusKey(CloseSquareKey);
                        break;
                    case '|':
                        WriteShiftPlusKey(BackslashKey);
                        break;
                    case ':':
                        WriteShiftPlusKey(SemicolonKey);
                        break;
                    case '"':
                        WriteShiftPlusKey(SingleQuoteKey);
                        break;
                    case '<':
                        WriteShiftPlusKey(CommaKey);
                        break;
                    case '>':
                        WriteShiftPlusKey(FullstopKey);
                        break;
                    case '?':
                        WriteShiftPlusKey(ForwardslashKey);
                        break;
                }
            }

            AuditCommand("WriteString", "" + stringToWrite);

        }

        /// <summary>
        /// Write a key
        /// </summary>
        /// <param name="key">Key as per definitions</param>
        public void WriteKey(byte key)
        {
            WriteKeyStroke(key);
            AuditCommand("WriteKey", "" + key);
        }

        /// <summary>
        /// Write a delay between events
        /// </summary>
        /// <param name="delay">Delay in ms</param>
        public void WriteDelay(int delay)
        {
            AuditCommand("WriteDelay", "" + delay);
            events.Add(new AutomationEvent(0, 0, 0, -1, delay, AutomationEventType.Keyboard));
            timer += delay;
        }

        /// <summary>
        /// Write a mouse move event
        /// </summary>
        /// <param name="mousePoint">Point to move to</param>
        public void WriteMouseMove(MousePoint mousePoint)
        {
            AuditCommand("WriteMouseMove", " x,y=" + mousePoint.X + "," + mousePoint.Y);
            events.Add(new AutomationEvent(MouseEventFlags.Move, mousePoint, mouseDelay, AutomationEventType.Mouse));
        }

        /// <summary>
        /// Write a mouse click event
        /// </summary>
        /// <param name="mouseEventFlags">Mouse events</param>
        public void WriteMouseClick(MouseEventFlags mouseEventFlags)
        {
            AuditCommand("WriteMouseClick", " mouseEventFlags =" + mouseEventFlags);
            events.Add(new AutomationEvent(mouseEventFlags, new MousePoint(), mouseDelay, AutomationEventType.Mouse));
        }

        /// <summary>
        /// Close up automation
        /// </summary>
        public void Close()
        {
            if (auditWriter != null)
            {
                auditWriter.Close();
            }
            if (logWriter != null)
            {
                logWriter.Close();
            }
        }

        /// <summary>
        /// Primitive vision - take a screenshot, do OCR and return text
        /// </summary>
        /// <returns>Entire text of screen</returns>
        public string Vision()
        {
            string imageFilename = CaptureScreen(0);
            string vision = GetStringFromImage(imageFilename).Result;
            AuditCommand("Vision", " filename='" + imageFilename + "' vision='" + vision + "'");
            return (vision);
        }

        /// <summary>
        /// Capture screen for OCR purposes
        /// </summary>
        /// <param name="limit">Brightness limit</param>
        /// <returns></returns>
        private string CaptureScreen(float limit)
        {
            string imageFilename = workingDirectory + processName.Replace(" ", "") + DateTime.Now.ToShortDateString().Replace("/", "") + Guid.NewGuid().ToString() + ".bmp";

            Image img = CaptureScreen();
            
            int width = img.Width * 2;
            int height = img.Height * 2;
            Bitmap bitmap = ResizeImage(img, width, height);

            if (limit != 0)
            {
                for (int i = 0; i < bitmap.Width; i++)
                {
                    for (int j = 0; j < bitmap.Height; j++)
                    {
                        Color c = bitmap.GetPixel(i, j);
                        if (c.GetBrightness() < limit)
                        {
                            bitmap.SetPixel(i, j, Color.Black);
                        }
                        else
                        {
                            bitmap.SetPixel(i, j, Color.White);
                        }
                    }
                }
            }

            bitmap.Save(imageFilename);

            string returnFilename = imageFilename;
            return (returnFilename);
        }

        /// <summary>
        /// Resize image
        /// </summary>
        /// <param name="image">Input image</param>
        /// <param name="width">New width</param>
        /// <param name="height">New height</param>
        /// <returns>Bitmap</returns>
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.AssumeLinear;
                graphics.InterpolationMode = InterpolationMode.Default;
                graphics.SmoothingMode = SmoothingMode.None;
                graphics.PixelOffsetMode = PixelOffsetMode.None;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        /// <summary>
        /// Wait to see specific text
        /// </summary>
        /// <param name="stringToSee">String to see</param>
        /// <param name="timeoutSecs">How many seconds to wait</param>
        /// <param name="errorMessage">Error message on timeout</param>
        public void WaitToSee(string stringToSee, int timeoutSecs, string errorMessage)
        {
            AuditCommand("WaitToSee", " stringToSee='" + stringToSee + "' timeoutSecs='" + timeoutSecs + "' errorMessage='" + errorMessage + "'");
            string vision = "";
            int timeout = 0;
            while (!vision.Contains(stringToSee))
            {
                timeout++;
                vision = Vision();
                Thread.Sleep(1000);
                if (timeout > timeoutSecs)
                {
                    throw new Exception(errorMessage);
                }
            }
        }

        /// <summary>
        /// Return specific location x,y of a word of text on a line
        /// </summary>
        /// <param name="searchLine">Filter on specific line</param>
        /// <param name="searchWord">Filter on specific word</param>
        /// <returns>Location</returns>
        public MousePoint ReturnWordLocation(string searchLine, string searchWord)
        {
            MousePoint point = new MousePoint();
            string imageFilename = CaptureScreen(0);
            point = ReturnWordLocation(imageFilename, searchLine, searchWord).Result;
            AuditCommand("ReturnWordLocation", " searchLine='" + searchLine + "' searchWord='" + searchWord + "' filename='" + imageFilename + "' location=" + point.X + "," + point.Y);
            return (point);
        }

        /// <summary>
        /// In some cases and color combinations you need to turn the image into black and white to find words
        /// </summary>
        /// <param name="searchWord">Search word</param>
        /// <param name="brightness">Float of brightness defining black/white</param>
        /// <returns>Location</returns>
        public MousePoint ReturnWordLocationAdvanced(string searchWord, float brightness)
        {
            MousePoint point = new MousePoint();
            string imageFilename = CaptureScreen(brightness);
            point = ReturnWordLocation(imageFilename, searchWord, searchWord).Result;
            AuditCommand("ReturnWordLocationAdvanced", " searchWord='" + searchWord + "' filename='" + imageFilename + "' location=" + point.X + "," + point.Y);
            return (point);
        }

        /// <summary>
        /// Return word location of word in an image
        /// </summary>
        /// <param name="filePath">Path to image</param>
        /// <param name="searchLine">Search line</param>
        /// <param name="searchWord">Search word</param>
        /// <returns>Location</returns>
        private async Task<MousePoint> ReturnWordLocation(string filePath, string searchLine, string searchWord)
        {
            MousePoint point = new MousePoint();
            var engine = OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language("en-US"));
            var file = await StorageFile.GetFileFromPathAsync(filePath).AsTask().ConfigureAwait(false);
            var stream = await file.OpenAsync(FileAccessMode.Read).AsTask().ConfigureAwait(false);
            var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream).AsTask().ConfigureAwait(false);
            var softwareBitmap = await decoder.GetSoftwareBitmapAsync().AsTask().ConfigureAwait(false);
            var ocrResult = await engine.RecognizeAsync(softwareBitmap).AsTask().ConfigureAwait(false);
            AuditCommand("ReturnWordLocation", "OCR result = '"+ocrResult.Text+"'");

            foreach (var line in ocrResult.Lines)
            {
                AuditCommand("ReturnWordLocation", "Line = '" + line.Text + "'");
                if (line.Text.Contains(searchLine))
                {
                    foreach (var word in line.Words)
                    {
                        AuditCommand("ReturnWordLocation", "Word = '" + word.Text + "'");
                        if (word.Text.Contains(searchWord))
                        {
                            point.X = (int) (word.BoundingRect.X/2) + 5;
                            point.Y = (int) (word.BoundingRect.Y/2) + 5;
                            break;
                        }
                    }
                }
            }
            return point;
        }

        /// <summary>
        /// Get full text from an image
        /// </summary>
        /// <param name="filePath">Path to image</param>
        /// <returns>Text</returns>
        private async Task<string> GetStringFromImage(string filePath)
        {
            var engine = OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language("en-US"));
            var file = await StorageFile.GetFileFromPathAsync(filePath).AsTask().ConfigureAwait(false);
            var stream = await file.OpenAsync(FileAccessMode.Read).AsTask().ConfigureAwait(false);
            var decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream).AsTask().ConfigureAwait(false);
            var softwareBitmap = await decoder.GetSoftwareBitmapAsync().AsTask().ConfigureAwait(false);
            var ocrResult = await engine.RecognizeAsync(softwareBitmap).AsTask().ConfigureAwait(false);            
            return ocrResult.Text;
        }

        /// <summary>
        /// Capture current desktop screen
        /// </summary>
        /// <returns>Image of desktop</returns>
        public Image CaptureScreen()
        {
            IntPtr desktop = GetDesktopWindow();
            return CaptureWindow(desktop);
        }

        /// <summary>
        /// Capture image of window
        /// </summary>
        /// <param name="handle">Handle to window</param>
        /// <returns>Image</returns>
        private Image CaptureWindow(IntPtr handle)
        {
            IntPtr hdcSrc = GetWindowDC(handle);
            RECT windowRect = new RECT();
            GetWindowRect(handle, ref windowRect);
            int width = windowRect.right - windowRect.left;
            int height = windowRect.bottom - windowRect.top;

            IntPtr hdcDest = CreateCompatibleDC(hdcSrc);
            IntPtr hBitmap = CreateCompatibleBitmap(hdcSrc, width, height);
            IntPtr hOld = SelectObject(hdcDest, hBitmap);
            BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, SRCCOPY);
            SelectObject(hdcDest, hOld);
            DeleteDC(hdcDest);
            ReleaseDC(handle, hdcSrc);
            Image img = Image.FromHbitmap(hBitmap);
            DeleteObject(hBitmap);

            return img;
        }

    }
}