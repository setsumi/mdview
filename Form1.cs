using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
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
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // ID for the Open item on the system menu
        private int SYSMENU_OPEN_ID = 0x1;

        private readonly string _exeDir;
        private readonly string _userdataFolder;
        private string _resultFile;
        private bool _initOnce = false;
        private bool _initForm = true;
        private FormWindowState _lastWindowState = FormWindowState.Normal;
        private readonly Plexiglass _plexiGlass;
        private readonly Random _random = new Random((int)DateTime.Now.Ticks & 0x0000FFFF);

        public FormMain(string[] args)
        {
            InitializeComponent();

            this.AllowDrop = true;
            _exeDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            _userdataFolder = Path.Combine(_exeDir, $"{Application.ProductName}_WebView2Cache");
            _plexiGlass = new Plexiglass(this);

            if (Properties.Settings.Default.winWidth > 0)
                this.Width = Properties.Settings.Default.winWidth;
            if (Properties.Settings.Default.winHeight > 0)
                this.Height = Properties.Settings.Default.winHeight;
            if (Properties.Settings.Default.winMaximized)
                this.WindowState = FormWindowState.Maximized;
            _lastWindowState = this.WindowState;

            Rectangle workingArea = Screen.FromControl(this).WorkingArea;
            RECT rect;
            GetWindowRect(this.Handle, out rect);
            int bottomOffset = workingArea.Bottom - 1 - rect.Bottom;
            if (bottomOffset < 0)
            {
                this.Top += bottomOffset;
                if (this.Top < workingArea.Top) this.Top = workingArea.Top;
            }

            if (!IsWebView2Installed())
            {
                MessageBox.Show("Microsoft Edge WebView2 is not installed.\n\nhttps://developer.microsoft.com/en-us/microsoft-edge/webview2/#download-section\n\n(Ctrl+C) - copy text",
                    this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Close();
            }

            _initForm = false;
        }

        private void FormMain_Shown(object sender, EventArgs e)
        {
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                Open(args[1]);
            }
            else
            {
                label1.Text = "Open file (Ctrl+O)...";
                webView21.Visible = false;
            }
        }

        public static bool IsWebView2Installed()
        {
            try
            {
                var version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                return version != null;
            }
            catch (WebView2RuntimeNotFoundException)
            {
                return false;
            }
        }

        private void Open(string filename)
        {
            label1.Text = "Loading...";
            this.Text = "LOADING...";
            if (webView21.Visible) _plexiGlass.Show();
            this.Refresh();
            Do(filename);
            this.Text = filename + "   - " + Application.ProductName;
            webView21.Visible = true;
            _plexiGlass.Hide();
            webView21.Focus();
        }

        async private void Do(string filename)
        {
            if (File.Exists(_resultFile)) File.Delete(_resultFile);
            _resultFile = Path.Combine(_exeDir, $"result{GenerateRandomString(15)}.html");

            // man format
            if (IsManExtension(Path.GetExtension(filename)))
            {
                string groff_file = Path.Combine(_exeDir, "groff", "bin", "groff.exe");
                if (!File.Exists(groff_file)) throw new FileNotFoundException($"File not found{Environment.NewLine}{groff_file}");
                ProcessStartInfo si = new ProcessStartInfo("cmd.exe");
                si.Arguments = $"/C \"{groff_file} -mandoc -Thtml {filename} > {_resultFile}\"";
                si.WindowStyle = ProcessWindowStyle.Hidden;
                using (var p = Process.Start(si)) p.WaitForExit();
            }
            else // markdown format
            {
                string result = @"<!DOCTYPE html><html><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width, initial-scale=1""></head><body>";
                string mdtext = File.ReadAllText(filename);
                var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
                result += Markdown.ToHtml(mdtext, pipeline);
                result += "<style>" + File.ReadAllText(Path.Combine(_exeDir, "style.css")) + "</style>";
                result += @"</body></html>";
                File.WriteAllText(_resultFile, result, Encoding.UTF8);
            }

            if (!_initOnce)
            {
                _initOnce = true;
                var env = await CoreWebView2Environment.CreateAsync(null, _userdataFolder, new CoreWebView2EnvironmentOptions("--disk-cache-dir=nul"));
                await webView21.EnsureCoreWebView2Async(env);
            }
            if (!File.Exists(_resultFile)) throw new FileNotFoundException($"File not found{Environment.NewLine}{_resultFile}");
            webView21.Source = new Uri(FormatFileURL(_resultFile));
        }

        private string GenerateRandomString(int length)
        {
            const string pool = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            var builder = new StringBuilder();
            for (var i = 0; i < length; i++)
            {
                var c = pool[_random.Next(0, pool.Length)];
                builder.Append(c);
            }
            return builder.ToString();
        }

        private string FormatFileURL(string filepath)
        {
            return @"file:///" + filepath.Replace(Path.DirectorySeparatorChar, '/');
        }

        private bool IsManExtension(string ext)
        {
            bool rv = false;
            if (string.IsNullOrEmpty(ext)) return false;

            for (int i = 1; i <= 8; i++)
            {
                if (ext == $".{i}")
                {
                    rv = true;
                    break;
                }
            }
            if (!rv)
            {
                rv = string.Compare(ext, ".mdoc", true) == 0 || string.Compare(ext, ".man", true) == 0;
            }
            return rv;
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
            _plexiGlass.Close();
            if (File.Exists(_resultFile)) File.Delete(_resultFile);
        }

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
            ofd.Filter = "All Files (*.*)|*.*|Markdown Files (*.md, *.markdown)|*.md;*.markdown|Unix Man Files (*.1, *.2, *.3, *.4, *.5, *.6, *.7, *.8, *.man, *.mdoc)|*.1;*.2;*.3;*.4;*.5;*.6;*.7;*.8;*.man;*.mdoc";
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
            AppendMenu(hSysMenu, MF_STRING, SYSMENU_OPEN_ID, "&Open file (Ctrl+O)...");
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

    public class Plexiglass : Form
    {
        private Form _tocover;
        public Plexiglass(Form tocover)
        {
            this.BackColor = Color.DarkGray;
            this.Opacity = 0.15;      // Tweak as desired
            this.FormBorderStyle = FormBorderStyle.None;
            this.ControlBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.AutoScaleMode = AutoScaleMode.None;
            this.Location = tocover.PointToScreen(Point.Empty);
            this.ClientSize = tocover.ClientSize;
            tocover.LocationChanged += Cover_LocationChanged;
            tocover.ClientSizeChanged += Cover_ClientSizeChanged;

            this.Owner = tocover;
            this.Visible = false;
            _tocover = tocover;
        }
        public new void Show()
        {
            base.Show(_tocover);
            _tocover.Focus();
        }
        private void Cover_LocationChanged(object sender, EventArgs e)
        {
            // Ensure the plexiglass follows the owner
            this.Location = this.Owner.PointToScreen(Point.Empty);
        }
        private void Cover_ClientSizeChanged(object sender, EventArgs e)
        {
            // Ensure the plexiglass keeps the owner covered
            this.ClientSize = this.Owner.ClientSize;
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Restore owner
            this.Owner.LocationChanged -= Cover_LocationChanged;
            this.Owner.ClientSizeChanged -= Cover_ClientSizeChanged;
            base.OnFormClosing(e);
        }
    }

}
