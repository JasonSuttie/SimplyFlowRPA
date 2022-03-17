using System;
using SimplyFlowWindows;

namespace WindowsAutomatorDemo
{
    class Program
    {
        public static void WordPadExample()
        {
            // Start process we want to work with
            WindowsAutomator windowsAutomator = new WindowsAutomator("CreateWordPadDocument", @"c:\automate\");
            windowsAutomator.LogMessage("WordPadExample","Starting wordpad process");
            windowsAutomator.StartProcess(@"wordpad.exe", "");

            // Ensure process is running before we carry on
            if (windowsAutomator.IsProcessRunning())
            {
                // Wait for Wordpad form to become visible
                windowsAutomator.WaitToSee("Document - WordPad", 20, "Wordpad didn't start correctly");

                // Write Text                       
                windowsAutomator.WriteString("Demo document content");
                windowsAutomator.Run();

                // Click and select
                var clickLocation = windowsAutomator.ReturnWordLocation("document", "document");
                windowsAutomator.WriteMouseMove(clickLocation);
                windowsAutomator.WriteMouseClick(WindowsAutomator.MouseEventFlags.LeftDown);
                windowsAutomator.WriteMouseClick(WindowsAutomator.MouseEventFlags.LeftUp);
                windowsAutomator.Run();

                // ALT F + x to exit
                windowsAutomator.WriteAltPlusKey("f");
                windowsAutomator.WriteString("x");
                windowsAutomator.WriteDelay(1000);
                windowsAutomator.Run();

                // Don't save
                var dontSaveLocation = windowsAutomator.ReturnWordLocationAdvanced("Don't", 0.7f);
                windowsAutomator.WriteDelay(1000);
                windowsAutomator.WriteMouseMove(dontSaveLocation);
                windowsAutomator.WriteMouseClick(WindowsAutomator.MouseEventFlags.LeftDown);
                windowsAutomator.WriteMouseClick(WindowsAutomator.MouseEventFlags.LeftUp);
                windowsAutomator.Run();

            }

            windowsAutomator.LogMessage("WordPadExample", "Finished wordpad process");
            windowsAutomator.Close();

        }

        [STAThread]
        public static void Main(string[] args)
        {
            WordPadExample();        
        }
    }
}