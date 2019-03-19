using DiscordRPC;
using DiscordRPC.Logging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Discord_Rich_Presence
{
    public partial class Form1 : Form
    {
        private string clientId = "557302945771683850";
        private string partyId = "ae488379-351d-4a4f-ad32-2b9b01c91657";
        private static int discordPipe = -1;

        private DiscordRpcClient client;
        private Timestamps timestamps;

        private Timer datetimeRefresh;

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", EntryPoint = "GetWindowText", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetTimeStamp();

            client = new DiscordRpcClient(clientId);
            client.Logger = new ConsoleLogger() { Level = LogLevel.Warning };
            client.Initialize();
            client.SetPresence(new RichPresence()
            {

                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                State = GetState(),
                Timestamps = timestamps,
                Assets = new Assets()
                {
                    LargeImageKey = "large_rin",
                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                    SmallImageKey = "small_offline",
                    SmallImageText = "未指定"
                }
            });

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;

            numericUpDown1.Enabled = checkBox3.Checked;
            numericUpDown2.Enabled = checkBox3.Checked;

            Timer timer = new Timer();
            timer.Tick += new EventHandler(TimerEvent);
            timer.Interval = 1000;
            timer.Start();

            if (checkBox2.Checked)
            {
                if (datetimeRefresh == null)
                {
                    datetimeRefresh = new Timer();
                    datetimeRefresh.Tick += new EventHandler(DateTimeEvent);
                    datetimeRefresh.Interval = 1000;
                    datetimeRefresh.Start();
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetTimeStamp();
            SetPresence(true);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetTimeStamp();
            SetPresence();
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            SetTimeStamp();
            SetPresence();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SetTimeStamp();
            SetPresence();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = false;
            checkBox2.Enabled = false;

            SetPresence();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = true;
            checkBox2.Enabled = true;

            SetTimeStamp();
            SetPresence();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = true;
            checkBox2.Enabled = true;

            SetTimeStamp();
            SetPresence();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                SetPresence();
            }
            else if (radioButton2.Checked)
            {
                if (radioButton2.Checked)
                    timestamps = Timestamps.Now;
                SetPresence();
            }
            else if (radioButton3.Checked)
            {
                DateTime dateTime = DateTime.Now;

                TimeSpan timeSpan1 = new TimeSpan(dateTimePicker1.Value.Hour, dateTimePicker1.Value.Minute, dateTimePicker1.Value.Second);
                TimeSpan timeSpan2 = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                TimeSpan timeSpan = timeSpan1 - timeSpan2;

                if (radioButton3.Checked)
                    timestamps = Timestamps.FromTimeSpan(timeSpan);
                SetPresence();
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                if (datetimeRefresh == null)
                {
                    datetimeRefresh = new Timer();
                    datetimeRefresh.Tick += new EventHandler(DateTimeEvent);
                    datetimeRefresh.Interval = 1000;
                    datetimeRefresh.Start();
                }
            }
            else
            {
                if (datetimeRefresh != null)
                {
                    datetimeRefresh.Stop();
                    datetimeRefresh = null;
                }
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown1.Enabled = checkBox3.Checked;
            numericUpDown2.Enabled = checkBox3.Checked;
        }

        private bool IsCorrectSize()
        {
            if (Decimal.ToInt32(numericUpDown1.Value) <= Decimal.ToInt32(numericUpDown2.Value))
                return true;
            else
                return false;
        }

        private void ShowErrorMessage()
        {
            MessageBox.Show("最大値よりも大きい値は指定できません。\r\n最大値と同じかそれ以下の値を指定してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string GetState()
        {
            if (checkBox1.Checked)
            {
                StringBuilder sb = new StringBuilder(65535);
                GetWindowText(GetForegroundWindow(), sb, 65535);

                string state1 = (textBox1.Text + " - " + sb.ToString()).Length < 121 ? (textBox1.Text + " - " + sb.ToString()) : (textBox1.Text + " - " + sb.ToString()).Substring(0, 120);
                string state2 = (sb.ToString().Length < 121 ? sb.ToString() : sb.ToString().Substring(0, 120));

                return (textBox1.Text.Length > 0 ? state1 : state2);
            }
            else
            {
                return (textBox1.Text.Length > 0 ? (textBox1.Text.Length < 121 ? textBox1.Text : textBox1.Text.Substring(0, 120)) : "");
            }
        }

        private void SetTimeStamp()
        {
            if (radioButton2.Checked)
            {
                timestamps = Timestamps.Now;
            }
            else if (radioButton3.Checked)
            {
                DateTime dateTime = DateTime.Now;

                TimeSpan timeSpan1 = new TimeSpan(dateTimePicker1.Value.Hour, dateTimePicker1.Value.Minute, dateTimePicker1.Value.Second);
                TimeSpan timeSpan2 = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                TimeSpan timeSpan = timeSpan1 - timeSpan2;

                timestamps = Timestamps.FromTimeSpan(timeSpan);
            }
        }

        private void SetPresence()
        {
            if (checkBox3.Checked)
            {
                if (!IsCorrectSize()) return;
                if (radioButton1.Checked)
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            break;
                    }
                }
                else
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            break;
                    }
                }
            }
            else
            {
                if (radioButton1.Checked)
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            break;
                    }
                }
                else
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            break;
                    }
                }
            }
        }

        private void SetPresence(bool b)
        {
            if (!b) SetPresence();

            if (checkBox3.Checked)
            {
                if (!IsCorrectSize()) return;
                if (radioButton1.Checked)
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            comboBox2.SelectedIndex = 0;
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            comboBox2.SelectedIndex = 11;
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            comboBox2.SelectedIndex = 17;
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            comboBox2.SelectedIndex = 22;
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                    }
                }
                else
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            comboBox2.SelectedIndex = 0;
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            comboBox2.SelectedIndex = 11;
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            comboBox2.SelectedIndex = 17;
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            comboBox2.SelectedIndex = 22;
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Party = new Party()
                                {
                                    ID = partyId,
                                    Size = Decimal.ToInt32(numericUpDown1.Value),
                                    Max = Decimal.ToInt32(numericUpDown2.Value)
                                },
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                    }
                }
            }
            else
            {
                if (radioButton1.Checked)
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            comboBox2.SelectedIndex = 0;
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            comboBox2.SelectedIndex = 11;
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            comboBox2.SelectedIndex = 17;
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            comboBox2.SelectedIndex = 22;
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                    }
                }
                else
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_online",
                                    SmallImageText = "オンライン"
                                }
                            });
                            comboBox2.SelectedIndex = 0;
                            break;
                        case 1:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_idle",
                                    SmallImageText = "退席中"
                                }
                            });
                            comboBox2.SelectedIndex = 11;
                            break;
                        case 2:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_dnd",
                                    SmallImageText = "取り込み中"
                                }
                            });
                            comboBox2.SelectedIndex = 17;
                            break;
                        case 3:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_streaming",
                                    SmallImageText = "配信中"
                                }
                            });
                            comboBox2.SelectedIndex = 22;
                            break;
                        case 4:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "オフライン"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                        default:
                            client.SetPresence(new RichPresence()
                            {
                                Details = (comboBox2.Text.Length > 0 ? comboBox2.Text : ""),
                                State = GetState(),
                                Timestamps = timestamps,
                                Assets = new Assets()
                                {
                                    LargeImageKey = "large_rin",
                                    LargeImageText = "しずりん☆ｶﾜ(・∀・)ｲｲ!!",
                                    SmallImageKey = "small_offline",
                                    SmallImageText = "未指定"
                                }
                            });
                            comboBox2.SelectedIndex = 29;
                            break;
                    }
                }
            }
        }

        private void TimerEvent(object sender, EventArgs e)
        {
            SetPresence();
        }

        private void DateTimeEvent(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
        }
    }
}
