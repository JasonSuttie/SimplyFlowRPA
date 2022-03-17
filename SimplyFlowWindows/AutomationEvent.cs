using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static SimplyFlowWindows.WindowsAutomator;

namespace SimplyFlowWindows
{
    class AutomationEvent
    {
        private WindowsAutomator.AutomationEventType type;

        public MouseEventFlags mouseEventFlags { get; private set; }
        public MousePoint mousePoint { get; private set; }

        public WindowsAutomator.AutomationEventType Type
        {
            get { return type; }
            set { type = value; }
        }

        private byte vk;

        public byte Vk
        {
            get { return vk; }
            set { vk = value; }
        }
        private byte scan;

        public byte Scan
        {
            get { return scan; }
            set { scan = value; }
        }
        private uint keyEventInfo;

        public uint KeyEventInfo
        {
            get { return keyEventInfo; }
            set { keyEventInfo = value; }
        }
        private int extre;

        public int Extra
        {
            get { return extre; }
            set { extre = value; }
        }

        private int delay;

        public int Delay
        {
            get { return delay; }
            set { delay = value; }
        }

        public AutomationEvent(byte vk, byte scan, uint keyInfo, int extra, int delay, WindowsAutomator.AutomationEventType type)
        {
            this.type = type;
            this.vk = vk;
            this.scan = scan;
            this.keyEventInfo = keyInfo;
            this.extre = extra;
            this.delay = delay;
            
        }

        public AutomationEvent(WindowsAutomator.MouseEventFlags mouseEventFlags, MousePoint mousePoint, int delay, WindowsAutomator.AutomationEventType type)
        {
            this.type = type;
            this.mouseEventFlags = mouseEventFlags;
            this.mousePoint = mousePoint;
            this.delay = delay;
        }
    }
}