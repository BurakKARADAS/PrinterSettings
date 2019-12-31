using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Printing;
using System.Net;
using System.Net.NetworkInformation;
using System.Management;
using DevExpress.XtraEditors;
using Microsoft.Win32;
namespace PrinterSettings
{
    public partial class Form1 : Form
    {
      
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = this.Text + " v" + Application.ProductVersion.ToString();
            var server = new PrintServer();
            var queues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            foreach (var queue in queues)
            {
               
                string printerName = queue.Name;
                string printerPort = queue.QueuePort.Name;
              

                if (!printerPort.Contains("192."))
                {
                    bool status = Connected(printerName);
                    GroupControl g = new GroupControl();
                    g.Size = new Size(this.Width-40, g.Height);
                    g.Dock = DockStyle.Top;

                    flowLayoutPanel1.Controls.Add(g);
                    FlowLayoutPanel f = new FlowLayoutPanel();
                    f.Dock = DockStyle.Fill;
                    f.FlowDirection = FlowDirection.TopDown;
                    g.Controls.Add(f);
                    Label l = new Label();
                    l.AutoSize = true;
                    l.Font = new Font("Tahoma", 13);
                    l.ImageAlign = ContentAlignment.MiddleLeft;
                    l.Text = "   Yazıcı Adı : " + printerName;
                    if (status) l.Image = Properties.Resources._1451487868_bullet_green;
                    else l.Image = Properties.Resources._1451487929_bullet_red;
                    f.Controls.Add(l);

                    Label l1 = new Label();
                    l1.AutoSize = true;
                    l1.Font = new Font("Tahoma", 13);
                    l1.Text = "   Port : " + printerPort;
                    f.Controls.Add(l1);
                }
                else
                {
                    bool status = Ping(printerPort);
                    GroupControl g = new GroupControl();
                    g.Size = new Size(this.Width - 40, g.Height);
                    g.Dock = DockStyle.Top;
                    flowLayoutPanel1.Controls.Add(g);
                    FlowLayoutPanel f = new FlowLayoutPanel();
                    f.Dock = DockStyle.Fill;
                    f.FlowDirection = FlowDirection.TopDown;
                    g.Controls.Add(f);
                    Label l = new Label();
                    l.AutoSize = true;
                    l.Font = new Font("Tahoma", 13);
                    l.ImageAlign = ContentAlignment.MiddleLeft;
                    l.Text = "   Yazıcı Adı : " + printerName;
                    if (status) l.Image = Properties.Resources._1451487868_bullet_green;
                    else l.Image = Properties.Resources._1451487929_bullet_red;
                    f.Controls.Add(l);

                    Label l1 = new Label();
                    l1.AutoSize = true;
                    l1.Font = new Font("Tahoma", 13);
                    l1.Text = "   Port : " + printerPort;
                    f.Controls.Add(l1);

                   
                }

            }

        }

        public bool Connected(string printer_name)
        {
            bool isConnected = false;

            try
            {


                ManagementScope scope = new ManagementScope(@"\root\cimv2");
                scope.Connect();


                ManagementObjectSearcher searcher = new
                 ManagementObjectSearcher("SELECT * FROM Win32_Printer");

                string printerName = "";
                foreach (ManagementObject printer in searcher.Get())
                {
                    printerName = printer["Name"].ToString();
                    if (printerName.Equals(printer_name))
                    {
                        if (printer["WorkOffline"].ToString().ToLower().Equals("true"))
                        {

                            isConnected = false;
                        }
                        else
                        {
                            isConnected = true;
                        }
                    }
                }
            }
            catch
            {
            }

            return isConnected;
        }
        public bool Ping(string p_IPAddress)
        {
            bool isConnected = false;
            try
            {
                IPAddress ip = IPAddress.Parse(p_IPAddress);

                Ping pingSender = new Ping();
                IPAddress address = IPAddress.Loopback;
                PingReply reply = pingSender.Send(p_IPAddress);

                if (reply.Status == IPStatus.Success)
                {
                    isConnected = true;
                }

            }
            catch
            {

            }
            return isConnected;
        }

        private void btnYenile_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();

            var server = new PrintServer();
            var queues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
            foreach (var queue in queues)
            {

                string printerName = queue.Name;
                string printerPort = queue.QueuePort.Name;


                if (!printerPort.Contains("192."))
                {
                    bool status = Connected(printerName);
                    GroupControl g = new GroupControl();
                    g.Size = new Size(this.Width - 40, g.Height);
                    g.Dock = DockStyle.Top;

                    flowLayoutPanel1.Controls.Add(g);
                    FlowLayoutPanel f = new FlowLayoutPanel();
                    f.Dock = DockStyle.Fill;
                    f.FlowDirection = FlowDirection.TopDown;
                    g.Controls.Add(f);
                    Label l = new Label();
                    l.AutoSize = true;
                    l.Font = new Font("Tahoma", 13);
                    l.ImageAlign = ContentAlignment.MiddleLeft;
                    l.Text = "   Yazıcı Adı : " + printerName;
                    if (status) l.Image = Properties.Resources._1451487868_bullet_green;
                    else l.Image = Properties.Resources._1451487929_bullet_red;
                    f.Controls.Add(l);

                    Label l1 = new Label();
                    l1.AutoSize = true;
                    l1.Font = new Font("Tahoma", 13);
                    l1.Text = "   Port : " + printerPort;
                    f.Controls.Add(l1);
                }
                else
                {
                    bool status = Ping(printerPort);
                    GroupControl g = new GroupControl();
                    g.Size = new Size(this.Width - 40, g.Height);
                    g.Dock = DockStyle.Top;
                    flowLayoutPanel1.Controls.Add(g);
                    FlowLayoutPanel f = new FlowLayoutPanel();
                    f.Dock = DockStyle.Fill;
                    f.FlowDirection = FlowDirection.TopDown;
                    g.Controls.Add(f);
                    Label l = new Label();
                    l.AutoSize = true;
                    l.Font = new Font("Tahoma", 13);
                    l.ImageAlign = ContentAlignment.MiddleLeft;
                    l.Text = "   Yazıcı Adı : " + printerName;
                    if (status) l.Image = Properties.Resources._1451487868_bullet_green;
                    else l.Image = Properties.Resources._1451487929_bullet_red;
                    f.Controls.Add(l);

                    Label l1 = new Label();
                    l1.AutoSize = true;
                    l1.Font = new Font("Tahoma", 13);
                    l1.Text = "   Port : " + printerPort;
                    f.Controls.Add(l1);


                }
            }
        }
    }
}
