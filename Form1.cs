using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Markdig;
using Microsoft.Web.WebView2.Core;

namespace mdview
{
    public partial class FormMain : Form
    {
        // P/Invoke constants
        private const int WM_SYSCOMMAND = 0x112;
        private const int MF_STRING = 0x0;
        private const int MF_SEPARATOR = 0x800;
        // P/Invoke declarations
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool AppendMenu(IntPtr hMenu, int uFlags, int uIDNewItem, string lpNewItem);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool InsertMenu(IntPtr hMenu, int uPosition, int uFlags, int uIDNewItem, string lpNewItem);
        // ID for the Open item on the system menu
        private int SYSMENU_OPEN_ID = 0x1;


        string _exeDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        bool _initOnce = false;
        bool _initForm = true;
        public FormMain(string[] args)
        {
            InitializeComponent();

            this.AllowDrop = true;

            if (Properties.Settings.Default.winWidth > 0)
                this.Width = Properties.Settings.Default.winWidth;
            if (Properties.Settings.Default.winHeight > 0)
                this.Height = Properties.Settings.Default.winHeight;
            if (Properties.Settings.Default.winMaximized)
                this.WindowState = FormWindowState.Maximized;

            if (args.Length > 0)
            {
                Open(args[0]);
            }
            else
            {
                label1.Text = "Open Markdown file (Ctrl+O)...";
                webView21.Visible = false;
            }
            _initForm = false;
        }

        private void Open(string filename)
        {
            this.Text = filename + " - " + Application.ProductName;
            label1.Text = "Loading...";
            this.Refresh();
            Do(filename);
            webView21.Visible = true;
            webView21.Focus();
        }

        async private void Do(string filename)
        {
            string mdtext = File.ReadAllText(filename);
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            string result = Markdown.ToHtml(mdtext, pipeline);
            result += "<style>" + File.ReadAllText(_exeDir + Path.DirectorySeparatorChar + "style.css") + "</style>";
            if (!_initOnce)
            {
                _initOnce = true;
                var env = await CoreWebView2Environment.CreateAsync(null, null, new CoreWebView2EnvironmentOptions("--disk-cache-dir=nul"));
                await webView21.EnsureCoreWebView2Async(env);
            }
            webView21.NavigateToString(result);
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        FormWindowState _lastWindowState = FormWindowState.Normal;
        private void FormMain_Resize(object sender, EventArgs e)
        {
            if (_initForm) return;

            if (WindowState != _lastWindowState)
            {
                _lastWindowState = WindowState;
                if (WindowState == FormWindowState.Maximized) // Maximized!
                {
                    Properties.Settings.Default.winMaximized = true;
                }
                else if (WindowState == FormWindowState.Normal)
                {
                    Properties.Settings.Default.winMaximized = false;
                }
            }
            else if (WindowState == FormWindowState.Normal) // resize
            {
                Properties.Settings.Default.winWidth = this.Width;
                Properties.Settings.Default.winHeight = this.Height;
            }
        }

        private void FormMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
            else if (e.KeyCode == Keys.O) // don't check if Ctrl is pressed since webview don't pass single letter key anyway
            {
                e.Handled = true;
                label1_MouseClick(null, null);
            }
        }

        private void webView21_KeyDown(object sender, KeyEventArgs e)
        {
            FormMain_KeyDown(sender, e);
        }

        private void label1_MouseClick(object sender, MouseEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Open(ofd.FileName);
            }
        }

        private void FormMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void FormMain_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            Open(files[0]);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            IntPtr hSysMenu = GetSystemMenu(this.Handle, false);
            AppendMenu(hSysMenu, MF_SEPARATOR, 0, string.Empty);
            AppendMenu(hSysMenu, MF_STRING, SYSMENU_OPEN_ID, "&Open Markdown file (Ctrl+O)...");
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if ((m.Msg == WM_SYSCOMMAND) && ((int)m.WParam == SYSMENU_OPEN_ID))
            {
                label1_MouseClick(null, null);
            }
        }
    }
}
