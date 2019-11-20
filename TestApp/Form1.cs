using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NikonFocusControl;

namespace TestApp
{
    public partial class Form1 : Form
    {
        FocusControl fc;

        public Form1()
        {
            InitializeComponent();
            Shown += Form1_Shown;
            FormClosing += Form1_FormClosing;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            fc.Dispose();
            fc = null;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            fc = new FocusControl();
            fc.DeviceConnected += DeviceConnected_handler;
            fc.DeviceDisconnected += DeviceDisconnected_handler;
        }

        private void DeviceConnected_handler(object sender, EventArgs e)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke((MethodInvoker)delegate
                {
                    DeviceConnected_handler(sender, e);
                });
                return;
            }

            //groupBox1.Enabled = true;
            BtnMove.Enabled = true;
            BtnConnect.Enabled = false;
            BtnDisconnect.Enabled = true;
        }

        private void DeviceDisconnected_handler(object sender, EventArgs e)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke((MethodInvoker)delegate
                {
                    DeviceDisconnected_handler(sender, e);
                });
                return;
            }

            //groupBox1.Enabled = false;
            BtnMove.Enabled = false;
            BtnConnect.Enabled = true;
            BtnDisconnect.Enabled = false;
        }

        private void BtnMove_Click(object sender, EventArgs e)
        {
            fc.Move((int)numericUpDown1.Value);
        }

        private void BtnMoveDisc_Click(object sender, EventArgs e)
        {
            try
            {
                fc.ConnectAndMove((int)numericUpDown1.Value);
            }
            catch (Exception exception)
            {
                //throw exception;
            }
        }

        private void BtnConnect_Click(object sender, EventArgs e)
        {
            fc.Connect();
            BtnConnect.Enabled = false;
        }

        private void BtnDisconnect_Click(object sender, EventArgs e)
        {
            fc.Disconnect();
            BtnDisconnect.Enabled = false;
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            numericUpDown1.Value = trackBar1.Value;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            trackBar1.Value = (int)numericUpDown1.Value;
        }
    }
}
