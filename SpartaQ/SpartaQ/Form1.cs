/* Copyright (C) SpartaQ, Inc - All Rights Reserved
 * SpartaQ tool is automated typing and mouse control program that onscreen image recognition.
 * Unauthorized copying of this file, via any medium is strictly prohibited
 * Proprietary and confidential
 * Written by Tatu-Pekka Saarinen <tatupeksaarinen@ymail.com>, 2019
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.IO;

namespace SpartaQ
{
    public partial class Form1 : Form
    {
        // DLL libraries used to manage hotkeys
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        const int MYACTION_HOTKEY_ID = 1;
        // DLL libraries used to manage mouse actions
        [DllImport("User32.dll", SetLastError = true)]
        public static extern int SendInput(int nInputs, ref INPUT pInputs, int cbSize);
        // DLL libraries to get cursor position
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out Point lpPoint);
        //CAPSLOCK
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags,
        UIntPtr dwExtraInfo);
        //mouse event constants
        const int MOUSEEVENTF_LEFTDOWN = 0x02;
        const int MOUSEEVENTF_LEFTUP = 0x04;
        const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        const int MOUSEEVENTF_RIGHTUP = 0x10;
        //input type constant
        const int INPUT_MOUSE = 0;
        static int choose2D = 0;
        static int[] choose2DArea;
        static string[,] KeyList = new string[255, 99];
        static int KeyLocation = 0;

        static bool running = false;
        static int speed = 500;
        static int iPhase = 0; //Value of active line
        static int iTask = 0; //Value of active task

        private void InitializeTimer()
        {
            // Call this procedure when the application starts.
            timer1.Interval = speed;
            timer1.Tick += new EventHandler(timer1_Tick);
            // Enable timer
            timer1.Enabled = true;
        }

        public void Combo()
        {
            Function.Items.AddRange(new object[] { "Message", "if", "else", "end", "ID", "TextCompare", "Parse", "Find_first", "Find_next", "Find_count", "Calculate", "Goto_Row", "Goto_ID", "Wait", "Write", "Turbo_Click", "Return_Start", "Write_Name", "Write_Small", "Write_Capital", "Write_Normal", "Copy", "Paste", "Mouse_Left", "Mouse_Right", "Mouse_Left_Down", "Mouse_Right_Down", "Mouse_Left_Up", "Mouse_Right_Up", "Mouse_X", "Mouse_Y", "Mouse_XAdd", "Mouse_YAdd", "Send_Key", "Capslock_On/Off", "AddCoordinates"});
        }

        private void timer1_Tick(object Sender, EventArgs e)
        {
            if (running)
            {
                if (iPhase < dataGridView1.Rows.Count - 1)
                {
                    timer1.Enabled = false;
                    string action = dataGridView1.Rows[iPhase].Cells[0].Value.ToString();
                    label1.Text = action;
                    switch (action)
                    {
                        case "Message":
                            MessageBox.Show(VariableParse());
                            ++iPhase;
                            break;

                        case "if":
                            double result = Convert.ToDouble(new DataTable().Compute(VariableParse(), null));
                            if (result == 1) { ++iPhase; }
                            else
                            {
                                int loop = 1;
                                for (int i = iPhase + 1; i < dataGridView1.RowCount; i++)
                                {
                                    string check = dataGridView1.Rows[i].Cells[0].Value.ToString();
                                    if ("end" == check) { --loop; }
                                    else if ("if" == check) { ++loop; }

                                    if (loop <= 0) { iPhase = i; i = dataGridView1.RowCount; }
                                }
                            }
                            break;

                        case "TextCompare":
                            string[] textComp = VariableParse().Split('=');
                            if (dataGridView1.Rows[iPhase].Cells[3].Value != null)
                            {
                                KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                                if (textComp[0] == textComp[1])
                                {
                                    KeyList[KeyLocation, 1] = "1";
                                }
                                else
                                {
                                    KeyList[KeyLocation, 1] = "0";
                                }
                                EraseDuplicate();
                            }
                            ++iPhase;
                            break;

                        case "end":
                            ++iPhase;
                            break;

                        case "Parse":
                            string PartP = VariableParse();
                            char[] firstPart = PartP.ToCharArray();
                            string lastPart = PartP.Substring(1, PartP.Length - 1);
                            string[] Presult = lastPart.Split(firstPart[0]);

                            if (dataGridView1.Rows[iPhase].Cells[3].Value != null)
                            {
                                KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();

                                for (int i = 0; i < Presult.Length; i++)
                                {
                                    KeyList[KeyLocation, i + 1] = Presult[i].ToString();
                                }
                                EraseDuplicate();
                            }
                            ++iPhase;
                            break;

                        case "Find_first":
                            ImageFind(0, 0, false, true, true);
                            ++iPhase;
                            break;

                        case "Find_count":
                            string countF = ImageFind(0, 0, true, true, true).ToString();
                            if (countF != "0")
                            {
                                KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                                KeyList[KeyLocation, 1] = countF;
                                KeyList[KeyLocation, 2] = countF;
                                EraseDuplicate();
                            }
                            ++iPhase;
                            break;

                        case "Find_next":
                            ImageFind(0, Cursor.Position.Y, false, true, true); //Eihän tää nyt toimi vitun pässi :O, kursori voi liikkua ilmankin find-ominaisuutta.
                            ++iPhase;
                            break;

                        case "Wait":
                            int variable;
                            if (dataGridView1.Rows[iPhase].Cells[1].Value != null)
                            {
                                System.Threading.Thread.Sleep(int.TryParse(dataGridView1.Rows[iPhase].Cells[1].Value.ToString(), out variable) ? variable : 0);
                            }
                            else
                            {
                                System.Threading.Thread.Sleep(speed);
                            }
                            ++iPhase;
                            break;

                        case "Write":
                            AutoTyper("");
                            ++iPhase;
                            break;

                        case "Return_Start":
                            iPhase = 0;
                            break;

                        case "Goto_Row":
                            iPhase = int.TryParse(dataGridView1.Rows[iPhase].Cells[1].Value.ToString(), out variable) ? variable : 0;
                            break;

                        case "Goto_ID":
                            int IDRow = 0;
                            string thisData = dataGridView1.Rows[iPhase].Cells[1].Value.ToString();
                            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                            {
                                string rowType = dataGridView1.Rows[IDRow].Cells[0].Value.ToString();
                                string rowData = dataGridView1.Rows[IDRow].Cells[1].Value.ToString();
                                if (rowData == thisData && rowType == "ID")
                                {
                                    iPhase = IDRow;
                                    break;
                                }
                                ++IDRow;
                            }
                            ++iPhase;
                            break;

                        case "Copy":
                            if (Clipboard.ContainsText(TextDataFormat.Text))
                            {
                                SendKeys.Send("^{c}");
                                if (dataGridView1.Rows[iPhase].Cells[3].Value != null)
                                {
                                    KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                                    KeyList[KeyLocation, 1] = Clipboard.GetText(TextDataFormat.Text);
                                    EraseDuplicate();
                                }
                            }
                            ++iPhase;
                            break;

                        case "Paste":
                            SendKeys.Send("^{v}");
                            ++iPhase;
                            break;

                        case "Turbo_Click":
                            TurboClicker();
                            ++iPhase;
                            break;

                        case "Write_Name":
                            AutoTyper("Name");
                            ++iPhase;
                            break;

                        case "Write_Normal":
                            AutoTyper("Basic");
                            ++iPhase;
                            break;

                        case "Write_Small":
                            AutoTyper("Small");
                            ++iPhase;
                            break;

                        case "Write_Capital":
                            AutoTyper("Capital");
                            ++iPhase;
                            break;

                        case "Mouse_Left":
                            Clicker("left_click");
                            ++iPhase;
                            break;

                        case "Mouse_Right":
                            Clicker("right_click");
                            ++iPhase;
                            break;

                        case "Mouse_Left_Down":
                            Clicker("left_down");
                            ++iPhase;
                            break;

                        case "Mouse_Right_Down":
                            Clicker("right_down");
                            ++iPhase;
                            break;

                        case "Mouse_Left_Up":
                            Clicker("left_up");
                            ++iPhase;
                            break;

                        case "Mouse_Right_Up":
                            Clicker("right_up");
                            ++iPhase;
                            break;

                        case "Mouse_X":
                            variable = int.TryParse(VariableParse(), out variable) ? variable : 0;
                            MouseMove(variable, Cursor.Position.Y);
                            ++iPhase;
                            break;

                        case "Mouse_Y":
                            variable = int.TryParse(VariableParse(), out variable) ? variable : 0;
                            MouseMove(Cursor.Position.X, variable);
                            ++iPhase;
                            break;

                        case "Mouse_XAdd":
                            variable = int.TryParse(VariableParse(), out variable) ? variable : 0;
                            MouseMove(Cursor.Position.X + variable, Cursor.Position.Y);
                            ++iPhase;
                            break;

                        case "Mouse_YAdd":
                            variable = int.TryParse(VariableParse(), out variable) ? variable : 0;
                            MouseMove(Cursor.Position.X, Cursor.Position.Y + variable);
                            ++iPhase;
                            break;

                        case "Calculate":
                            result = Convert.ToDouble(new DataTable().Compute(VariableParse(), null));
                            if (dataGridView1.Rows[iPhase].Cells[3].Value != null)
                            {
                                KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                                KeyList[KeyLocation, 1] = result.ToString();
                                EraseDuplicate();
                            }
                            ++iPhase;
                            break;

                        case "Send_Key":
                            string sKey = dataGridView1.Rows[iPhase].Cells[1].Value.ToString();
                            if (sKey != null)
                            {
                                SendKeys.Send(sKey);
                            }
                            ++iPhase;
                            break;

                        case "Capslock_On/Off":
                            bool DoIt = false;
                            string OnOrOff;
                            if (dataGridView1.Rows[iPhase].Cells[1].Value != null)
                            {
                                OnOrOff = dataGridView1.Rows[iPhase].Cells[1].Value.ToString();
                                if (OnOrOff == "On")
                                {
                                    if (!Control.IsKeyLocked(Keys.CapsLock)) // Checks Capslock is on
                                    {
                                        DoIt = true;
                                    }
                                }

                                else if (OnOrOff == "Off")
                                {
                                    if (Control.IsKeyLocked(Keys.CapsLock)) // Checks Capslock is off
                                    {
                                        DoIt = true;
                                    }
                                }
                            }

                            else
                            {
                                DoIt = true;
                            }

                            if (DoIt)
                            {
                                const int KEYEVENTF_EXTENDEDKEY = 0x1;
                                const int KEYEVENTF_KEYUP = 0x2;
                                keybd_event(0x14, 0x45, KEYEVENTF_EXTENDEDKEY, (UIntPtr)0);
                                keybd_event(0x14, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                                (UIntPtr)0);
                            }
                            ++iPhase;
                            break;

                        case "AddCoordinates":
                            result = Convert.ToDouble(new DataTable().Compute(VariableParse(), null));
                            if (dataGridView1.Rows[iPhase].Cells[3].Value != null)
                            {
                                KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                                int i;
                                for (i = 0; KeyList[KeyLocation, i] != null; i++){}
                                KeyList[KeyLocation, i + 1] = Cursor.Position.X.ToString();
                                KeyList[KeyLocation, i + 2] = Cursor.Position.Y.ToString();

                                EraseDuplicate();
                            }
                            ++iPhase;
                            break;
                             
                        default:
                            label1.Text = "Fail!";
                            ++iPhase;
                            break;
                    }
                    timer1.Interval = speed;
                    timer1.Enabled = true;
                }

                else
                {
                    //label1.Text = "Stop";
                    running = false;
                    timer1.Enabled = false;
                    iPhase = 0;
                    Array.Clear(KeyList, 0, KeyList.Length);
                    KeyLocation = 0;
                    if (iTask < dataGridView2.Rows.Count)
                    {
                        dataGridView2.Rows[iTask].Cells[2].Value = true;
                        ++iTask;
                        RunTime();
                    }
                }
            }
        }

        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        public struct INPUT
        {
            public uint type;
            public MOUSEINPUT mi;
        };

        //
        //Form
        //

        public Form1()
        {
            choose2DArea = new int[] { 0, 0, 0, 0 };

            // Modifier combination keys codes: Alt = 1, Ctrl = 2, Shift = 4, Win = 8
            // ALT+CTRL = 1 + 2 = 3 , CTRL+SHIFT = 2 + 4 = 6...

            InitializeComponent();
            label1.BackColor = System.Drawing.Color.Transparent;
            label2.BackColor = System.Drawing.Color.Transparent;
            label3.BackColor = System.Drawing.Color.Transparent;
            label4.BackColor = System.Drawing.Color.Transparent;
            label5.BackColor = System.Drawing.Color.Transparent;
            label6.BackColor = System.Drawing.Color.Transparent;
            Combo();
            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //GLobal KEY-used
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)
            {
                if (m.WParam.ToInt32() == MYACTION_HOTKEY_ID)
                {
                    if (choose2D > 0)
                    {
                        Point cursorPos;
                        GetCursorPos(out cursorPos);

                        if (choose2D == 1)
                        {
                            choose2DArea[0] = cursorPos.X;
                            choose2DArea[1] = cursorPos.Y;
                            label1.Text = cursorPos.X + "," + cursorPos.Y;
                            choose2D = 2;
                        }

                        else if (choose2D == 2)
                        {
                            if (choose2DArea[0] < cursorPos.X && choose2DArea[1] < cursorPos.Y)
                            {
                                choose2DArea[2] = cursorPos.X;
                                choose2DArea[3] = cursorPos.Y;
                                label1.Text = choose2DArea[0] + "," + choose2DArea[1] + "," + choose2DArea[2] + "," + choose2DArea[3];
                                ScreenPart(choose2DArea);
                            }
                            choose2DArea = new int[] { 0, 0, 0, 0 };
                            choose2D = 0;
                        }
                    }

                    else
                    {
                        label1.Text = "Stop";
                        iPhase = 0;
                        running = false;
                        timer1.Enabled = false;
                        Array.Clear(KeyList, 0, KeyList.Length);
                        KeyLocation = 0;
                    }
                }
            }
            base.WndProc(ref m);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RunTime();
        }

        private void RunTime()
        {
            if (iTask < dataGridView2.Rows.Count - 1)
            {
                if (dataGridView2.Rows[iTask].Cells[2].Value == null)
                {
                    KeyList[KeyLocation, 0] = "TaskName";
                    KeyList[KeyLocation, 1] = dataGridView2.Rows[iTask].Cells[0].Value.ToString();
                    KeyList[KeyLocation, 2] = dataGridView2.Rows[iTask].Cells[1].Value.ToString();
                    ++KeyLocation;
                    running = true;
                    timer1.Enabled = true;
                    WindowState = FormWindowState.Minimized;
                }

                else
                {
                    if ((bool)dataGridView2.Rows[iTask].Cells[2].Value == true)
                    {
                        ++iTask;
                    }
                    else
                    {
                        KeyList[KeyLocation, 0] = "TaskName";
                        KeyList[KeyLocation, 1] = dataGridView2.Rows[iTask].Cells[0].Value.ToString();
                        KeyList[KeyLocation, 2] = dataGridView2.Rows[iTask].Cells[1].Value.ToString();
                        ++KeyLocation;
                        running = true;
                        timer1.Enabled = true;
                        WindowState = FormWindowState.Minimized;
                    }
                }
            }

            else
            {
                MessageBox.Show("All tasks completed!");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        //Speed
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            running = false;
            if (int.TryParse(textBox1.Text, out int value))
            {
                if (value != 0)
                {
                    speed = Convert.ToInt32(textBox1.Text);
                    timer1.Interval = speed;
                }
            }

            else
            {
                timer1.Interval = 1000;
                textBox1.Text = "1000";
            }
        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            running = false;
            Application.Exit();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                DataGridViewImageCell cell = (DataGridViewImageCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (choose2D == 0)
                {
                    choose2D = 1;
                }
            }
            GridColor();
        }

        public void ScreenPart(int[] imageSize)
        {
            //Create a new bitmap.
            var bmpScreenshot = new Bitmap(imageSize[2] - imageSize[0], imageSize[3] - imageSize[1], PixelFormat.Format32bppArgb);
            var gfxScreenshot = Graphics.FromImage(bmpScreenshot);

            // Take the screenshot from the upper left corner to the right bottom corner.
            gfxScreenshot.CopyFromScreen(imageSize[0],
                                        imageSize[1],
                                        Screen.PrimaryScreen.Bounds.X,
                                        Screen.PrimaryScreen.Bounds.Y,
                                        Screen.PrimaryScreen.WorkingArea.Size,
                                        CopyPixelOperation.SourceCopy);
            label1.Text = Screen.PrimaryScreen.Bounds.Size.ToString();

            // Make the default transparent color transparent for myBitmap.
            //bmpScreenshot.MakeTransparent(bmpScreenshot.GetPixel(0, 0));

            // Save the screenshot to the specified path that the user has chosen.
            bmpScreenshot.Save("Screenshot.png", ImageFormat.Png);

            //dataGridView1.Rows.Add("Copy", "Hello", bmpScreenshot);
            dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].SetValues("Find_first", "", bmpScreenshot);
            GridColor();
        }


        public void dataGridView1_ValueChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value == null)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Green;
                }

                else
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }
            }
        }

        public void MouseMove(int x, int y)
        {
            //this will hold the location where to click
            Point clickLocation = new Point(x, y);
            //set cursor position to memorized location
            Cursor.Position = clickLocation;
            //SendKeys.Send("{S}"); < this is just for save
        }

        public void Clicker(string action)
        {
            //set up the INPUT struct and fill it for the mouse down
            INPUT i = new INPUT();
            i.type = INPUT_MOUSE;
            i.mi.dx = 0;
            i.mi.dy = 0;

            if (action == "left_down" || action == "left_click") { i.mi.dwFlags = MOUSEEVENTF_LEFTDOWN; }
            else if (action == "right_down" || action == "right_click") { i.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN; }

            i.mi.dwExtraInfo = IntPtr.Zero;
            i.mi.mouseData = 0;
            i.mi.time = 0;
            //send the input 
            SendInput(1, ref i, Marshal.SizeOf(i));
            //set the INPUT for mouse up and send it
            if (action == "left_up" || action == "left_click") { i.mi.dwFlags = MOUSEEVENTF_LEFTUP; }
            else if (action == "right_up" || action == "right_click") { i.mi.dwFlags = MOUSEEVENTF_RIGHTUP; }
            SendInput(1, ref i, Marshal.SizeOf(i));
        }

        //Adds row numbers to DataGrid
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var rowHeaderText = (e.RowIndex + 1).ToString();
            var dgv = sender as DataGridView;
            using (SolidBrush brush = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                var textFormat = new StringFormat()
                {
                    Alignment = StringAlignment.Far,
                    LineAlignment = StringAlignment.Far
                };

                var bounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, dgv.RowHeadersWidth, e.RowBounds.Height);
                e.Graphics.DrawString(rowHeaderText, this.Font, brush, bounds, textFormat);
            }
        }

        private void AutoTyper(string format)
        {
            if (dataGridView1.Rows[iPhase].Cells[1].Value != null)
            {
                string symbols = VariableParse();
                string tranform = "";

                for (int i = 0; i < symbols.Length; i++)
                {
                    string letter = symbols.Substring(i, 1);
                    if (format == "Small") { letter = letter.ToLower(); }
                    if (format == "Capital") { letter = letter.ToUpper(); }

                    if (format == "Name")
                    {
                        if (i == 0)
                        {
                            letter = letter.ToUpper();
                        }
                        else
                        {
                            letter = letter.ToLower();
                        }
                    }

                    else if (format == "Basic")
                    {
                        switch (letter)
                        {
                            case "Ä": letter = "A"; break;
                            case "Ö": letter = "O"; break;
                            case "Å": letter = "O"; break;
                            case "Ü": letter = "U"; break;
                            case "Ÿ": letter = "Y"; break;
                            case "Ï": letter = "I"; break;
                            case "Û": letter = "U"; break;
                            case "Î": letter = "I"; break;

                            case "ä": letter = "a"; break;
                            case "ö": letter = "o"; break;
                            case "å": letter = "o"; break;
                            case "ü": letter = "u"; break;
                            case "ÿ": letter = "y"; break;
                            case "ï": letter = "i"; break;
                        }
                    }
                    else { }
                    SendKeys.Send(letter);
                    tranform = tranform + letter;

                }

                if (dataGridView1.Rows[iPhase].Cells[3].Value != null)
                {
                    KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                    KeyList[KeyLocation, 1] = tranform.ToString();
                    EraseDuplicate();
                }
            }
        }

        private string VariableParse() //Parses the text for variables or "keywords". These keywords are paired with stored values in array.
        {
            string newLine = "";
            if (dataGridView1.Rows[iPhase].Cells[1].Value != null)
            {
                string line = dataGridView1.Rows[iPhase].Cells[1].Value.ToString();
                int varPoint = 0;
                for (int i = 0; i < line.Length; i++)
                {
                    string letter = line.Substring(i, 1);
                    if (letter == "(" && varPoint == 0) { ++varPoint; }//All the text that are inside round brackets are handled as keywords.
                    if (varPoint > 0)
                    {
                        if (letter == ")")
                        {
                            string line2 = line.Substring(i - varPoint + 2, varPoint - 2);
                            label1.Text = line2;
                            int subkey = 0;
                            string[] parts = line2.Split('.');//Numbers after dots are for array pointer.

                            for (int k = parts.Length-1; k >= 0; k--)//There can be sub-keyword that you can make loop with.
                            {
                                for (int j = 0; KeyList[j, 0] != null; j++)
                                {
                                    if (KeyList[j, 0] == parts[k])// The zero place at array is for keyword-name. 1 is usually for x coordinate values and 2 is for y.
                                    {
                                        //FIX THIS! at the moment this only works with maximum of one sub-keyword (keyw.keyw.numb). Better make the last part (.numb) as if statement.
                                        if (k > 0 && parts.Length > 2 && KeyList[j, 0] != null) //Checks if sub-keyword. && parts.Length > 1 && KeyList[j, 1] != null // Convert.ToInt32(parts[k + 1]) 
                                        {
                                            subkey = Convert.ToInt32(KeyList[j, Convert.ToInt32(parts[k + 1])]);
                                        }

                                        else if (k == 0 && parts.Length > 2 && KeyList[j, subkey] != null) //The main keyword
                                        {
                                            newLine = newLine + KeyList[j, subkey];
                                        }

                                        else if (k == 0 && parts.Length > 1 && KeyList[j, Convert.ToInt32(parts[k + 1])] != null) //The main keyword
                                        {
                                            newLine = newLine + KeyList[j, Convert.ToInt32(parts[k + 1])];
                                        }

                                        else { newLine = newLine + KeyList[j, 0]; } //If variable only has it's name and nothing stored.
                                        break;
                                    }
                                }
                            }
                            varPoint = 0;
                        }
                        else { ++varPoint; }
                    }
                    else { newLine = newLine + letter; }
                }
            }
            else
            {
                newLine = "1";
            }
            return newLine;
        }

        private void TurboClicker()
        {
            //set up the INPUT struct and fill it for the mouse down
            INPUT i = new INPUT();
            i.type = INPUT_MOUSE;
            i.mi.dx = 0;
            i.mi.dy = 0;

            int counter = int.TryParse(dataGridView1.Rows[iPhase].Cells[1].Value.ToString(), out counter) ? counter : 0;

            for (int j = 0; j < counter; j++)
            {
                i.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
                i.mi.dwExtraInfo = IntPtr.Zero;
                i.mi.mouseData = 0;
                i.mi.time = 0;
                //send the input 
                SendInput(1, ref i, Marshal.SizeOf(i));
                i.mi.dwFlags = MOUSEEVENTF_LEFTUP;
                SendInput(1, ref i, Marshal.SizeOf(i));
            }
        }

        //Array Key duplicate deleting function.

        public void EraseDuplicate()
        {
            for (int i = 0; i < KeyLocation; i++)
            {
                if (KeyList[i, 0] == KeyList[KeyLocation, 0])
                {
                    for (int j = 0; j < KeyList.GetLength(1); j++)
                    {
                        KeyList[i, j] = KeyList[KeyLocation, j];
                        KeyList[KeyLocation, j] = null;
                    }

                    i = KeyLocation;

                    if (KeyLocation > 0)
                    {
                        --KeyLocation;
                    }
                    break;
                }
            }
            ++KeyLocation;
        }

        //Image Finding Algorithm
        /////////////////////////

        public int ImageFind(int pointX, int pointY, bool count, bool center, bool rotation)
        {
            string parameters;

            if (dataGridView1.Rows[iPhase].Cells[1].Value.ToString() == null || dataGridView1.Rows[iPhase].Cells[1].Value.ToString() == "")
            {
                parameters = "100,100";
            }
            else
            {
                parameters = dataGridView1.Rows[iPhase].Cells[1].Value.ToString();
            }

            string[] sParam = parameters.Split(','); //accuracy%, tolerance%
            int acCount = 100;
            int toleCount = 100;
            bool acIsNumber = int.TryParse(sParam[0], out acCount);
            bool toleIsNumber = int.TryParse(sParam[1], out toleCount);
            Bitmap image1 = (Bitmap)dataGridView1.Rows[iPhase].Cells[2].Value;
            int allPixels = image1.Height * image1.Width;
            // Make the default transparent color transparent for myBitmap.
            int CheckX = 0;
            int CheckY = 0;
            int counter = 0;

            // Loop through the images pixels to reset color.

            for (int y = pointY; y < image1.Height; y++)
            {
                for (int x = pointX; x < image1.Width; x++)
                {
                    if (image1.GetPixel(x, y) != image1.GetPixel(0, 0))
                    {
                        CheckX = x;
                        CheckY = y;
                        label1.Text = "Anomality: " + x + "," + y;
                        break;
                    }
                }
            }

            // Create a new bitmap. Again.
            var bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb);
            var gfxScreenshot = Graphics.FromImage(bmpScreenshot);

            // Take the screenshot from the upper left corner to the right bottom corner.
            gfxScreenshot.CopyFromScreen(0,
                                        0,
                                        Screen.PrimaryScreen.Bounds.X,
                                        Screen.PrimaryScreen.Bounds.Y,
                                        Screen.PrimaryScreen.WorkingArea.Size,
                                        CopyPixelOperation.SourceCopy);

            // Loop through the images pixels to reset color.
            //label1.Text = "Screen shot done";

            //Save first vertical line of Find_images pixels in the list to later check rotation.
            Color rcol2 = image1.GetPixel(0, image1.Height-1);
            Color[] rcol = new Color[image1.Height];

            if (rotation)
            {
                pointX = pointX + image1.Width;
                pointY = pointY + image1.Height;
                for (int i = 0; i < image1.Height; i++)
                {
                    rcol[i] = image1.GetPixel(0, image1.Height-1-i);
                }
            }

            for (int y = pointY+200; y < bmpScreenshot.Height - image1.Height; y++)
            {
                for (int x = pointX; x < bmpScreenshot.Width - image1.Width; x++)
                {
                    //MouseMove(100, y);
                    //Clicker("left_click");
                    Color col = rcol[0];//image1.GetPixel(CheckX, CheckY);
                    Color SCol = bmpScreenshot.GetPixel(x + CheckX, y + CheckY);
                    int checkAC = 0;
                    int checkTOLE = 0;
                    float tolePlus = 0.05f;
                    float toleMinus = 0.95f;

                    //(col.B + (255 - col.B) * 0.2 > SCol.B && col.B * 0.8 < SCol.B) && (col.R + (255 - col.R) * 0.2 > SCol.R && col.R * 0.8 < SCol.R) && (col.G + (255 - col.G) * 0.2 > SCol.G && col.G * 0.8 < SCol.G)
                    if ((col.B + (255 - col.B) * tolePlus >= SCol.B && col.B * toleMinus <= SCol.B) && (col.R + (255 - col.R) * tolePlus >= SCol.R && col.R * toleMinus <= SCol.R) && (col.G + (255 - col.G) * tolePlus >= SCol.G && col.G * toleMinus <= SCol.G))
                    {
                        //MouseMove(x, y + image1.Height-1);
                        //Clicker("left_click");
                        //tolePlus =  1f-(float)acCount/100f;
                        //toleMinus = (float)acCount /100f;
                        tolePlus = 0.2f;
                        toleMinus = 0.8f;

                        if (rotation)
                        {
                            int jtimes = 0;
                            int allaround = 0;
                            int together = 0;
                            double angle;

                            for (double j = 0; j < 360; j++)
                            {
                                angle = Math.PI * j / 180.0;
                                int skipper = 0;

                                for (int i = 0; i < image1.Height; i++)
                                {
                                    Color rotateCol = bmpScreenshot.GetPixel(Convert.ToInt32(x + Math.Cos(angle) * i), Convert.ToInt32(y + image1.Height-1 + Math.Sin(angle) * i));

                                    if (i >= image1.Height - 1)
                                    {
                                        //MouseMove(Convert.ToInt32(x + Math.Cos(angle) * i), Convert.ToInt32(y + image1.Height + Math.Sin(angle) * i));
                                        //Clicker("right_click");
                                        ++jtimes;
                                        allaround = allaround + Convert.ToInt32(j);
                                        together = allaround / jtimes;
                                        //j = 360;
                                    }

                                    if ((rcol[i].B + (255 - rcol[i].B) * tolePlus >= rotateCol.B && rcol[i].B * toleMinus <= rotateCol.B) && (rcol[i].R + (255 - rcol[i].R) * tolePlus >= rotateCol.R && rcol[i].R * toleMinus <= rotateCol.R) && (rcol[i].G + (255 - rcol[i].G) * tolePlus >= rotateCol.G && rcol[i].G * toleMinus <= rotateCol.G))
                                    {
                                        //MouseMove(Convert.ToInt32(x + Math.Cos(angle) * i), Convert.ToInt32(y + image1.Height + Math.Sin(angle) * i));
                                        //Clicker("left_click");
                                    }

                                    else
                                    {
                                        ++skipper;
                                        if (skipper > image1.Height/3) //skipper > (float)image1.Height*(1-toleCount/100
                                        {
                                            i = image1.Height;
                                            skipper = 0;
                                        }
                                    }
                                }
                            }

                            if (together > 0)
                            {
                                label1.Text = "Oll Korrect: " + x + "," + y;
                                if (dataGridView1.Rows[iPhase].Cells[3].Value != null && !count)
                                {
                                    KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                                    KeyList[KeyLocation, 1] = (x).ToString();
                                    KeyList[KeyLocation, 2] = (y + image1.Height).ToString();
                                    KeyList[KeyLocation, 3] = together.ToString();
                                    EraseDuplicate();
                                }
                                else { MouseMove(x, y + image1.Height); }

                                if (!count)
                                {
                                    x = bmpScreenshot.Width;
                                    y = bmpScreenshot.Height;
                                }

                                else
                                {
                                    ++counter;
                                }
                            }
                        }

                        if (x != bmpScreenshot.Width && y != bmpScreenshot.Height)
                        {
                            for (int i = 0; i < image1.Width; i++)
                            {
                                for (int j = 0; j < image1.Height; j++)
                                {
                                    if (i == image1.Width - 1 && j == image1.Height - 1)
                                    {
                                        label1.Text = "Oll Korrect: " + x + "," + y;
                                        if (dataGridView1.Rows[iPhase].Cells[3].Value != null && !count)
                                        {
                                            KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                                            KeyList[KeyLocation, 1] = (x + (i / 2)).ToString();
                                            KeyList[KeyLocation, 2] = (y + (j / 2)).ToString();
                                            EraseDuplicate();
                                        }
                                        else { MouseMove(x + (i / 2), y + (j / 2)); }

                                        if (!count)
                                        {
                                            x = bmpScreenshot.Width;
                                            y = bmpScreenshot.Height;
                                        }

                                        else
                                        {
                                            ++counter;
                                        }

                                        i = image1.Width;
                                        j = image1.Height;
                                    }

                                    else
                                    {
                                        Color ImgPix = image1.GetPixel(i, j);
                                        Color ScreenPix = bmpScreenshot.GetPixel(x + i, y + j);
                                        float tole = (1 - (float)toleCount / 100) * 255;

                                        if (ImgPix.A != 0)
                                        {
                                            if (ImgPix.R - tole > ScreenPix.R || ImgPix.R + tole < ScreenPix.R || ImgPix.G - tole > ScreenPix.G || ImgPix.G + tole < ScreenPix.G || ImgPix.B - tole > ScreenPix.B || ImgPix.B + tole < ScreenPix.B) //|| ImgPix.G - tole <= ScreenPix.G || ImgPix.G + tole >= ScreenPix.G || ImgPix.B - tole <= ScreenPix.B || ImgPix.B + tole >= ScreenPix.B
                                            {
                                                ++checkAC;
                                                if ((float)checkAC / (float)allPixels >= 1 - (float)acCount / 100)
                                                {
                                                    i = image1.Width;
                                                    j = image1.Height;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (dataGridView1.Rows[iPhase].Cells[3].Value != null)
            {
                if (KeyLocation > 0)
                {
                    if (KeyList[KeyLocation - 1, 0] != dataGridView1.Rows[iPhase].Cells[3].Value.ToString())
                    {
                        KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                        KeyList[KeyLocation, 1] = "-1";
                        KeyList[KeyLocation, 2] = "-1";
                        EraseDuplicate();
                    }
                }
                else
                {
                    if (KeyList[KeyLocation, 0] == null)
                    {
                        KeyList[KeyLocation, 0] = dataGridView1.Rows[iPhase].Cells[3].Value.ToString();
                        KeyList[KeyLocation, 1] = "-1";
                        KeyList[KeyLocation, 2] = "-1";
                        EraseDuplicate();
                    }
                }
            }

            return counter;
        }

        public void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add();
            for (int i = dataGridView1.RowCount - 1; i > dataGridView1.CurrentCell.RowIndex; i--)
            {
                dataGridView1.Rows[i].SetValues(dataGridView1.Rows[i - 1].Cells[0].Value, dataGridView1.Rows[i - 1].Cells[1].Value, dataGridView1.Rows[i - 1].Cells[2].Value, dataGridView1.Rows[i - 1].Cells[3].Value);
            }
            dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].SetValues();
            GridColor();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int i = dataGridView1.CurrentCell.RowIndex;
            if (i > 0)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[dataGridView1.RowCount - 1].SetValues(dataGridView1.Rows[i - 1].Cells[0].Value, dataGridView1.Rows[i - 1].Cells[1].Value, dataGridView1.Rows[i - 1].Cells[2].Value, dataGridView1.Rows[i - 1].Cells[3].Value);
                dataGridView1.Rows[i - 1].SetValues(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value);
                dataGridView1.Rows[i].SetValues(dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value, dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[1].Value, dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[2].Value, dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[3].Value);
                int j = dataGridView1.RowCount;
                dataGridView1.CurrentCell = dataGridView1[0, j - 2];
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);
                dataGridView1.CurrentCell = dataGridView1[0, i - 1];
                dataGridView1.Rows[dataGridView1.RowCount - 1].SetValues(null, null, null, null);
            }
            GridColor();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int i = dataGridView1.CurrentCell.RowIndex;
            if (i < dataGridView1.RowCount - 2)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[dataGridView1.RowCount - 1].SetValues(dataGridView1.Rows[i + 1].Cells[0].Value, dataGridView1.Rows[i + 1].Cells[1].Value, dataGridView1.Rows[i + 1].Cells[2].Value, dataGridView1.Rows[i + 1].Cells[3].Value);
                dataGridView1.Rows[i + 1].SetValues(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value);
                dataGridView1.Rows[i].SetValues(dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value, dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[1].Value, dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[2].Value, dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[3].Value);
                int j = dataGridView1.RowCount;
                dataGridView1.CurrentCell = dataGridView1[0, j - 2];
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);
                dataGridView1.CurrentCell = dataGridView1[0, i + 1];
                dataGridView1.Rows[dataGridView1.RowCount - 1].SetValues(null, null, null, null);
            }
            GridColor();
        }

        //Save
        private void button6_Click(object sender, EventArgs e)
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory + textBox2.Text + @"\";
            // If directory does not exist, this creates it. 
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            TextWriter writer = new StreamWriter(directory + @"Text.txt"); //new StreamWriter(@"C:\Users\Tatu\Desktop\Text.txt");

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (j != 2)
                        {
                            writer.Write(dataGridView1.Rows[i].Cells[j].Value.ToString());
                        }

                        else
                        {
                            writer.Write(i);
                            Bitmap image1 = (Bitmap)dataGridView1.Rows[i].Cells[j].Value;
                            image1.Save(directory + i + ".png", ImageFormat.Png);
                        }
                    }
                    writer.Write("|");
                }
                writer.WriteLine("");
            }
            writer.Close();
            label1.Text = "Saved: " + directory;
        }

        //Load
        private void button7_Click(object sender, EventArgs e)
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory + textBox2.Text + @"\";
            // If directory does not exist, this creates it. 
            if (Directory.Exists(directory))
            {
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();
                //TextReader reader = new StreamReader(directory + @"Text.txt");
                string[] lines = File.ReadAllLines(directory + @"Text.txt");
                for (int i = 0; i < lines.Length; i++)
                {
                    dataGridView1.Rows.Add();
                    string[] substring = lines[i].Split('|');
                    for (int j = 0; j < substring.Length - 1; j++)
                    {
                        if (substring[j] != "")
                        {
                            if (j != 2)
                            {
                                dataGridView1.Rows[i].Cells[j].Value = substring[j];
                            }

                            else
                            {
                                if (File.Exists(directory + i + ".png"))
                                {
                                    FileStream fs = new FileStream(directory + i + ".png", FileMode.OpenOrCreate);
                                    dataGridView1.Rows[i].Cells[j].Value = new Bitmap(fs);
                                    fs.Dispose();
                                }
                            }

                        }
                    }
                }
                label1.Text = "Loaded";
                GridColor();
            }
            else
            {
                label1.Text = "Invalid project name";
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            iTask = dataGridView2.CurrentCell.RowIndex;
        }

        private void GridColor()
        {
            int iRow = 0;
            int isif = 0;
            bool colored = false;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (iRow < dataGridView1.Rows.Count - 1)
                {
                    string action = dataGridView1.Rows[iRow].Cells[0].Value.ToString();
                    switch (action)
                    {
                        case "Wait":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 155, 225);
                            colored = true;
                            break;

                        case "Goto_Row":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 100);
                            colored = true;
                            break;

                        case "Goto_ID":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 100);
                            colored = true;
                            break;

                        case "ID":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 100);
                            colored = true;
                            break;

                        case "if":
                            ++isif;
                            break;

                        case "end":
                            row.DefaultCellStyle.BackColor = Color.FromArgb(50 + (175 / isif), 100 + (125 / isif), 255);
                            colored = true;
                            --isif;
                            break;
                    }

                    if (colored == false)
                    {
                        if (isif > 0)
                        {
                            row.DefaultCellStyle.BackColor = Color.FromArgb(50 + (175 / isif), 100 + (125 / isif), 255);
                        }
                        else
                        {
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 255);
                        }
                    }

                    colored = false;
                    ++iRow;
                }
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {
            label6.BackColor = System.Drawing.Color.Transparent;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text.ToString() != "Ins")
            {
                RegisterHotKey(this.Handle, MYACTION_HOTKEY_ID, 0, (int)Keys.Insert);
            }

            else
            {
                RegisterHotKey(this.Handle, MYACTION_HOTKEY_ID, 0, (int)Keys.Escape);
            }
        }
    }
}
