using Database.Models;
using System;
using System.Drawing;
using System.Windows.Forms;
using Win32APIs;

namespace MultiRemoteDesktopClient
{
    public delegate void Connected(object sender, EventArgs e, int ListIndex);
    public delegate void Connecting(object sender, EventArgs e, int ListIndex);
    public delegate void LoginComplete(object sender, EventArgs e, int ListIndex);
    public delegate void Disconnected(object sender, AxMSTSCLib.IMsTscAxEvents_OnDisconnectedEvent e, int ListIndex);
    public delegate void OnFormClosing(object sender, FormClosingEventArgs e, int ListIndex, IntPtr Handle);
    public delegate void OnFormActivated(object sender, EventArgs e, int ListIndex, IntPtr Handle);
    public delegate void OnFormShown(object sender, EventArgs e, int ListIndex, IntPtr Handle);
    public delegate void ServerSettingsChanged(object sender, Model_ServerDetails sd, int ListIndex);

    public partial class RdpClientWindow : Form
    {
        public event Connected Connected;
        public event Connecting Connecting;
        public event LoginComplete LoginComplete;
        public event Disconnected Disconnected;
        public event OnFormClosing OnFormClosing;
        public event OnFormActivated OnFormActivated;
        public event OnFormShown OnFormShown;
        public event ServerSettingsChanged ServerSettingsChanged;

        public Model_ServerDetails _sd;

        // used to easly locate in Server lists (RemoteDesktopClient)
        private int _listIndex = 0;

        private bool _isFitToWindow = false;

        public RdpClientWindow(Model_ServerDetails sd, Form parent)
        {
            InitializeComponent();
            InitializeControl(sd);
            InitializeControlEvents();

            this.MdiParent = parent;
            this.Visible = true;
        }

        public void InitializeControl(Model_ServerDetails sd)
        {
            GlobalHelper.infoWin.AddControl(new object[] {
                btnFitToScreen
            });

            this._sd = sd;

            // Log connection attempt details
            Console.WriteLine("=== RDP Connection Initialization ===");
            Console.WriteLine($"Server: {sd.Server}");
            Console.WriteLine($"Username: {sd.Username}");
            Console.WriteLine($"Domain: {sd.Domain ?? "(none)"}");
            Console.WriteLine($"Password Length: {sd.Password?.Length ?? 0}");
            Console.WriteLine($"Port: {sd.Port}");

            rdpClient.Server = sd.Server;

            // Set domain and username separately
            if (!string.IsNullOrEmpty(sd.Domain))
            {
                rdpClient.Domain = sd.Domain;
                Console.WriteLine($"Domain set to: {sd.Domain}");
            }

            rdpClient.UserName = sd.Username;
            Console.WriteLine($"Username set to: {sd.Username}");

            rdpClient.AdvancedSettings2.ClearTextPassword = sd.Password;
            rdpClient.ColorDepth = sd.ColorDepth;
            rdpClient.DesktopWidth = sd.DesktopWidth;
            rdpClient.DesktopHeight = sd.DesktopHeight;
            rdpClient.FullScreen = sd.Fullscreen;
            

            // this fixes the rdp control locking issue
            // when lossing its focus
            //rdpClient.AdvancedSettings3.ContainerHandledFullScreen = -1;
            //rdpClient.AdvancedSettings3.DisplayConnectionBar = true;
            //rdpClient.FullScreen = true;
            //rdpClient.AdvancedSettings3.SmartSizing = true;
            //rdpClient.AdvancedSettings3.PerformanceFlags = 0x00000100;

            //rdpClient.AdvancedSettings2.allowBackgroundInput = -1;
            rdpClient.AdvancedSettings2.AcceleratorPassthrough = -1;
            rdpClient.AdvancedSettings2.Compress = -1;
            rdpClient.AdvancedSettings2.BitmapPersistence = -1;
            rdpClient.AdvancedSettings2.BitmapPeristence = -1;
            //rdpClient.AdvancedSettings2.BitmapCacheSize = 512;
            rdpClient.AdvancedSettings2.CachePersistenceActive = -1;

            // Match RDCMan settings exactly:
            // - "Warn if authentication fails" = CHECKED (AuthenticationLevel = 1)
            // - "Enable CredSSP support" = NOT CHECKED (EnableCredSspSupport = false)
            try
            {
                // Cast to AdvancedSettings5 or higher to access EnableCredSspSupport
                var advancedSettings5 = rdpClient.AdvancedSettings2 as dynamic;
                if (advancedSettings5 != null)
                {
                    advancedSettings5.AuthenticationLevel = 2; // Set to 2 (connect and don't warn if authentication fails)
                    advancedSettings5.EnableCredSspSupport = false;  // Disable CredSSP like RDCMan

                    // Add more settings that might help with error 516
                    // These settings can help prevent internal errors related to control initialization
                    advancedSettings5.EnableAutoReconnect = true;
                    advancedSettings5.MaxReconnectAttempts = 3;

                    // Try setting NegotiateSecurityLayer - important for non-CredSSP connections
                    advancedSettings5.NegotiateSecurityLayer = true;

                    // Allow connection even if authentication fails at client
                    advancedSettings5.AllowBackgroundInput = 0;

                    Console.WriteLine("Authentication Level: 2 (Connect and don't warn if authentication fails)");
                    Console.WriteLine("CredSSP Support Enabled: False");
                    Console.WriteLine("EnableAutoReconnect: True");
                    Console.WriteLine("NegotiateSecurityLayer: True");
                    Console.WriteLine("AllowBackgroundInput: 0");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Warning: Could not configure authentication: " + ex.Message);
            }

            // custom port
            if (sd.Port != 0)
            {
                rdpClient.AdvancedSettings2.RDPPort = sd.Port;
            }

            btnConnect.Enabled = false;

            panel1.Visible = false;
            tmrSC.Enabled = false;
        }

        public void InitializeControlEvents()
        {
            this.Shown += new EventHandler(RdpClientWindow_Shown);
            this.FormClosing += new FormClosingEventHandler(RdpClientWindow_FormClosing);

            btnDisconnect.Click += new EventHandler(ToolbarButtons_Click);
            btnConnect.Click += new EventHandler(ToolbarButtons_Click);
            btnReconnect.Click += new EventHandler(ToolbarButtons_Click);
            btnSettings.Click += new EventHandler(ToolbarButtons_Click);
            btnFullscreen.Click += new EventHandler(ToolbarButtons_Click);
            m_FTS_FitToScreen.Click += new EventHandler(ToolbarButtons_Click);
            m_FTS_Strech.Click += new EventHandler(ToolbarButtons_Click);
            btnPopout_in.Click += new EventHandler(btnPopout_in_Click);

            this.rdpClient.OnConnecting += new EventHandler(rdpClient_OnConnecting);
            this.rdpClient.OnConnected += new EventHandler(rdpClient_OnConnected);
            this.rdpClient.OnLoginComplete += new EventHandler(rdpClient_OnLoginComplete);
            this.rdpClient.OnDisconnected += new AxMSTSCLib.IMsTscAxEvents_OnDisconnectedEventHandler(rdpClient_OnDisconnected);
            this.rdpClient.OnWarning += new AxMSTSCLib.IMsTscAxEvents_OnWarningEventHandler(rdpClient_OnWarning);
            this.rdpClient.OnFatalError += new AxMSTSCLib.IMsTscAxEvents_OnFatalErrorEventHandler(rdpClient_OnFatalError);

            btnSndKey_TaskManager.Click += new EventHandler(SendKeys_Button_Click);

            tmrSC.Tick += new EventHandler(tmrSC_Tick);
        }

        public const int WM_LEAVING_FULLSCREEN = 0x4ff;
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x21)  // mouse click
            {
                this.rdpClient.Focus();
            }
            else if (m.Msg == WM_LEAVING_FULLSCREEN)
            {
            }

            base.WndProc(ref m);
        }

        PopupMDIContainer popupmdi = null;
        void btnPopout_in_Click(object sender, EventArgs e)
        {
            // we just can't move our entire form
            // into different window because of the ActiveX error
            // crying out about the Windowless control.

            if (int.Parse(btnPopout_in.Tag.ToString()) == 0)
            {
                popupmdi = new PopupMDIContainer();
                popupmdi.Show();
                popupmdi.PopIn(ref rdpPanelBase, this, this._sd.ServerName);

                btnPopout_in.Image = Properties.Resources.pop_in_16;
                btnPopout_in.Tag = 1;
            }
            else if (int.Parse(btnPopout_in.Tag.ToString()) == 1)
            {
                popupmdi.PopOut(ref rdpPanelBase, this);

                btnPopout_in.Image = Properties.Resources.pop_out_16;
                btnPopout_in.Tag = 0;
            }
        }

        #region EVENT: Send Keys
        void SendKeys_Button_Click(object sender, EventArgs e)
        {
            rdpClient.Focus();

            if (sender == btnSndKey_TaskManager)
            {
                //SendKeys.Send("(^%)");
                SendKeys.Send("(^%{END})");
            }

            //rdpClient.AdvancedSettings2.HotKeyCtrlAltDel;
        }
        #endregion

        #region EVENT: RDP Client

        void rdpClient_OnDisconnected(object sender, AxMSTSCLib.IMsTscAxEvents_OnDisconnectedEvent e)
        {
            Status("Disconnected from " + this._sd.Server);

            btnConnect.Enabled = true;
            btnDisconnect.Enabled = false;

            // Log disconnect reason
            Console.WriteLine("=== RDP Disconnected ===");
            Console.WriteLine($"OnDisconnected - Connected Status: {rdpClient.Connected}");
            Console.WriteLine($"Disconnect Reason Code: {e.discReason}");

            // Decode disconnect reason
            string reasonText = GetDisconnectReason(e.discReason);
            Console.WriteLine($"Disconnect Reason: {reasonText}");

            // Show message box for authentication failures
            if (e.discReason == 2308 || e.discReason == 264 || e.discReason == 1286 || e.discReason == 2055)
            {
                MessageBox.Show($"Connection failed: {reasonText}\n\nReason Code: {e.discReason}\n\nServer: {this._sd.Server}\nUsername: {this._sd.Username}\nDomain: {this._sd.Domain ?? "(none)"}",
                    "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // Show info for error 516 (internal error)
            else if (e.discReason == 516)
            {
                MessageBox.Show($"Connection failed with internal error.\n\nThis may be caused by:\n- Incompatible security settings\n- Display/resolution configuration issues\n- Server policy restrictions\n\nReason Code: {e.discReason}\n\nServer: {this._sd.Server}\nUsername: {this._sd.Username}\nDomain: {this._sd.Domain ?? "(none)"}\n\nTry checking the console output for more details.",
                    "Internal RDP Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (Disconnected != null)
            {
                Disconnected(this, e, this._listIndex);
            }
        }

        private string GetDisconnectReason(int reason)
        {
            // Common RDP disconnect reason codes
            switch (reason)
            {
                case 0: return "No error";
                case 1: return "Local disconnection";
                case 2: return "Remote disconnection by user";
                case 3: return "Remote disconnection by server / User initiated disconnect";
                case 260: return "DNS name lookup failure";
                case 262: return "Out of memory";
                case 264: return "Connection timed out";
                case 516: return "Internal error";
                case 518: return "Out of memory";
                case 520: return "Host not found";
                case 772: return "Winsock error";
                case 1030: return "Security error";
                case 1032: return "Encryption error";
                case 1286: return "License protocol error";
                case 2308: return "The specified computer name contains invalid characters";
                case 2055: return "Internal security error";
                case 2056: return "Internal security error";
                case 2822: return "Logon failure: unknown username or bad password";
                case 2825: return "Account restriction prevents logon";
                case 3079: return "Connection to remote PC lost";
                default: return $"Unknown disconnect reason (code: {reason})";
            }
        }

        void rdpClient_OnLoginComplete(object sender, EventArgs e)
        {
            Status("Loged in using " + this._sd.Username + " user account");

            { // check connection status on output
                Console.WriteLine("OnLoginComplete - Connected Status: " + rdpClient.Connected);
            }

            if (LoginComplete != null)
            {
                LoginComplete(this, e, this._listIndex);
            }
        }

        void rdpClient_OnConnected(object sender, EventArgs e)
        {
            Status("Connected to " + this._sd.Server);

            { // check connection status on output
                Console.WriteLine("OnConnected - Connected Status: " + rdpClient.Connected);
            }

            if (Connected != null)
            {
                Connected(this, e, this._listIndex);
            }
        }

        void rdpClient_OnConnecting(object sender, EventArgs e)
        {
            Status("Connecting to " + this._sd.Server);

            btnConnect.Enabled = false;
            btnDisconnect.Enabled = true;

            { // check connection status on output
                Console.WriteLine("OnConnecting - Connected Status: " + rdpClient.Connected);
            }

            if (Connecting != null)
            {
                Connecting(this, e, this._listIndex);
            }
        }

        void rdpClient_OnWarning(object sender, AxMSTSCLib.IMsTscAxEvents_OnWarningEvent e)
        {
            Console.WriteLine("=== RDP Warning ===");
            Console.WriteLine($"Warning Code: {e.warningCode}");

            string warningText = GetWarningDescription(e.warningCode);
            Console.WriteLine($"Warning: {warningText}");
        }

        void rdpClient_OnFatalError(object sender, AxMSTSCLib.IMsTscAxEvents_OnFatalErrorEvent e)
        {
            Console.WriteLine("=== RDP Fatal Error ===");
            Console.WriteLine($"Error Code: {e.errorCode}");

            string errorText = GetErrorDescription(e.errorCode);
            Console.WriteLine($"Fatal Error: {errorText}");

            MessageBox.Show($"Fatal RDP Error: {errorText}\n\nError Code: {e.errorCode}\n\nServer: {this._sd.Server}\nUsername: {this._sd.Username}\nDomain: {this._sd.Domain ?? "(none)"}",
                "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string GetWarningDescription(int warningCode)
        {
            switch (warningCode)
            {
                case 1: return "Certificate warning";
                case 2: return "Certificate name mismatch";
                case 3: return "Certificate expired";
                default: return $"Unknown warning (code: {warningCode})";
            }
        }

        private string GetErrorDescription(int errorCode)
        {
            switch (errorCode)
            {
                case 0: return "Internal error";
                case 1: return "Protocol error";
                case 2: return "Out of memory";
                case 3: return "Control error";
                case 4: return "Invalid parameter";
                default: return $"Unknown error (code: {errorCode})";
            }
        }

        #endregion

        #region EVENT: server settings window

        Rectangle ssw_GetClientWindowSize()
        {
            return rdpClient.RectangleToScreen(rdpClient.ClientRectangle);
        }

        void ssw_ApplySettings(object sender, Model_ServerDetails sd)
        {
            this._sd = sd;

            MessageBox.Show("This will restart your connection", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            Reconnect(true, false, false);
        }

        #endregion

        #region EVENT: other form controls

        void tmrSC_Tick(object sender, EventArgs e)
        {
            pictureBox1.BackgroundImage = GetCurrentScreen();
        }

        void ToolbarButtons_Click(object sender, EventArgs e)
        {
            if (sender == btnDisconnect)
            {
                Disconnect();
            }
            else if (sender == btnConnect)
            {
                Connect();
            }
            else if (sender == btnReconnect)
            {
                Reconnect(false, this._isFitToWindow, false);
            }
            else if (sender == btnSettings)
            {
                ServerSettingsWindow ssw = new ServerSettingsWindow(this._sd);

                ssw.ApplySettings += new ApplySettings(ssw_ApplySettings);
                ssw.GetClientWindowSize += new GetClientWindowSize(ssw_GetClientWindowSize);
                ssw.ShowDialog();

                this._sd = ssw.CurrentServerSettings();

                if (ServerSettingsChanged != null)
                {
                    ServerSettingsChanged(sender, this._sd, this._listIndex);
                }
            }
            else if (sender == btnFullscreen)
            {
                DialogResult dr = MessageBox.Show("You are about to enter in Fullscreen mode.\r\nBy default, the remote desktop resolution will be the same as what you see on the window.\r\n\r\nWould you like to resize it automatically based on your screen resolution though it will be permanent as soon as you leave in Fullscreen.\r\n\r\nNote: This will reconnect.", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)
                {
                    Reconnect(false, false, true);
                }
                else
                {
                    rdpClient.FullScreen = true;
                }
            }
            else if (sender == m_FTS_FitToScreen)
            {
                DialogResult dr = MessageBox.Show("This will resize the server resolution based on this current client window size, though it will not affect you current settings.\r\n\r\nDo you want to continue?", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

                if (dr == DialogResult.OK)
                {
                    Reconnect(true, true, false);
                }
            }
            else if (sender == m_FTS_Strech)
            {
                if (int.Parse(m_FTS_Strech.Tag.ToString()) == 0)
                {
                    rdpClient.AdvancedSettings3.SmartSizing = true;
                    m_FTS_Strech.Text = "Don't Stretch";
                    m_FTS_Strech.Tag = 1;
                }
                else
                {
                    rdpClient.AdvancedSettings3.SmartSizing = false;
                    m_FTS_Strech.Text = "Stretch";
                    m_FTS_Strech.Tag = 0;
                }
            }
        }

        void RdpClientWindow_Shown(object sender, EventArgs e)
        {
            if (OnFormShown != null)
            {
                OnFormShown(this, e, this._listIndex, this.Handle);
            }

            // stretch RD view
            ToolbarButtons_Click(m_FTS_Strech, null);
        }

        void RdpClientWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are you sure you want to close this window?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (dr == DialogResult.Yes)
            {
                Disconnect();
                rdpClient.Dispose();

                if (OnFormClosing != null)
                {
                    OnFormClosing(this, e, this._listIndex, this.Handle);
                }

                Dispose();
            }
            else
            {
                e.Cancel = true;
            }
        }

        void RdpClientWindow_Activated(object sender, EventArgs e)
        {
            this.rdpClient.Focus();

            if (OnFormActivated != null)
            {
                OnFormActivated(this, e, this._listIndex, this.Handle);
            }
        }

        #endregion

        #region METHOD: s

        public void Connect()
        {
            Status("Starting ...");
            rdpClient.Connect();
        }

        public void Disconnect()
        {
            Status("Disconnecting ...");
            rdpClient.DisconnectedText = "Disconnected";

            if (rdpClient.Connected != 0)
            {
                rdpClient.Disconnect();
            }
        }

        public void Reconnect(bool hasChanges, bool isFitToWindow, bool isFullscreen)
        {
            Disconnect();

            Status("Waiting for the server to properly disconnect ...");

            // wait for the server to properly disconnect
            while (rdpClient.Connected != 0)
            {
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();
            }

            Status("Reconnecting ...");

            if (hasChanges)
            {
                rdpClient.Server = this._sd.Server;
                rdpClient.UserName = this._sd.Username;
                rdpClient.AdvancedSettings2.ClearTextPassword = this._sd.Password;
                rdpClient.ColorDepth = this._sd.ColorDepth;

                this._isFitToWindow = isFitToWindow;

                if (isFitToWindow)
                {
                    rdpClient.DesktopWidth = this.rdpClient.Width;
                    rdpClient.DesktopHeight = this.rdpClient.Height;
                }
                else
                {
                    rdpClient.DesktopWidth = this._sd.DesktopWidth;
                    rdpClient.DesktopHeight = this._sd.DesktopHeight;
                }

                rdpClient.FullScreen = this._sd.Fullscreen;
            }

            if (isFullscreen)
            {
                rdpClient.DesktopWidth = Screen.PrimaryScreen.Bounds.Width;
                rdpClient.DesktopHeight = Screen.PrimaryScreen.Bounds.Height;

                rdpClient.FullScreen = true;
            }

            Connect();
        }

        public Image GetCurrentScreen()
        {
            return APIs.ControlToImage.GetControlScreenshot(this.panel2);
        }

        private void Status(string stat)
        {
            lblStatus.Text = stat;
        }

        #endregion

        #region PROPERTY

        public int ListIndex
        {
            get
            {
                return this._listIndex;
            }
            set
            {
                this._listIndex = value;
            }
        }

        #endregion
    }
}