using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace 第一頁
{
    public partial class NavigationBar : UserControl
    {
        public TabControl MainTabControl { get; set; }
        public Button BtnLoginNav { get; private set; }    // 未登入用按鈕
        public Label LblWelcomeText { get; private set; }  // 歡迎文字
        private PictureBox picUserIcon;                    // 登入後圖片按鈕
        private ContextMenuStrip userMenu;                 // 下拉選單
        private string currentRole = "User";               // 預設角色

        public event EventHandler LogoutClicked;
        public event EventHandler MemberCenterClicked;
        public event EventHandler AdminPanelClicked;
        public event EventHandler AppointmentClicked;
        public event EventHandler AppointmentAdminClicked;
        public NavigationBar()
        {
            InitializeComponent();

            // 1. 基本設定
            this.Dock = DockStyle.Top;
            this.Height = 50;
            this.BackColor = Color.Transparent;

            // 2. 佈局：左右兩欄
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));   // 左側伸縮
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));       // 右側自動

            // 3. 左側按鈕群
            var centerPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10, 10, 10, 5),
                AutoSize = true,
                WrapContents = false
            };
            centerPanel.Controls.Add(CreateButton("首頁", () => SelectTab("tabPageHome")));
            centerPanel.Controls.Add(CreateButton("AI塔羅占卜", () => SelectTab("tabPageAITarot")));
            centerPanel.Controls.Add(CreateButton("是否塔羅占卜", () => SelectTab("tabPageYesNoTarot")));
            centerPanel.Controls.Add(CreateButton("塔羅運勢", () => SelectTab("tabPageTimeTarot")));
            centerPanel.Controls.Add(CreateButton("預約占卜", () => SelectTab("tabPageAppointment")));

            // 4. 右側歡迎文字 + 按鈕
            var rightPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 10, 10, 5),
                AutoSize = true,
                WrapContents = false,
                Anchor = AnchorStyles.Right
            };

            // 4.1 歡迎文字
            LblWelcomeText = new Label
            {
                AutoSize = true,
                Text = "登入會員享有更多功能",
                Font = new Font("微軟正黑體", 9, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(5, 11, 5, 0)
            };
            rightPanel.Controls.Add(LblWelcomeText);

            // 4.2 未登入按鈕
            BtnLoginNav = CreateButton("登入", () => SelectTab("tabPageLogin"));
            BtnLoginNav.Width = 100;
            BtnLoginNav.Height = 30;
            rightPanel.Controls.Add(BtnLoginNav);

            // 4.3 登入後的圓形圖片按鈕
            picUserIcon = new PictureBox
            {
                Size = new Size(36, 36),
                SizeMode = PictureBoxSizeMode.Zoom,
                Cursor = Cursors.Hand,
                Visible = false,
                Margin = new Padding(5, 0, 5, 0)
            };
            picUserIcon.Click += PicUserIcon_Click;
            picUserIcon.Resize += (s, e) => SetCircularRegion(picUserIcon);
            rightPanel.Controls.Add(picUserIcon);

            // 5. 將左右面板加入主佈局
            layout.Controls.Add(centerPanel, 0, 0);
            layout.Controls.Add(rightPanel, 1, 0);
            this.Controls.Add(layout);

            // 6. 準備下拉選單
            InitUserMenu();
        }
        private void InitUserMenu(string role = "User")
        {
            userMenu = new ContextMenuStrip();
            userMenu.Items.Clear();
            if (role == "User")
            {
                userMenu.Items.Add("會員中心", null, (s, e) => MemberCenterClicked?.Invoke(this, EventArgs.Empty));
                userMenu.Items.Add("預約占卜", null, (s, e) => AppointmentClicked?.Invoke(this, EventArgs.Empty));
                userMenu.Items.Add("登出", null, (s, e) => LogoutClicked?.Invoke(this, EventArgs.Empty));
            }
            else
            {
                userMenu.Items.Add("管理者後台", null, (s, e) => AdminPanelClicked?.Invoke(this, EventArgs.Empty));
                userMenu.Items.Add("預約後台", null, (s, e) => AppointmentAdminClicked?.Invoke(this, EventArgs.Empty));
                userMenu.Items.Add("登出", null, (s, e) => LogoutClicked?.Invoke(this, EventArgs.Empty));
            }
        }

        private void PicUserIcon_Click(object sender, EventArgs e)
        {
            if (userMenu != null && picUserIcon.Visible)
            {
                // 顯示在圖片下方中央
                var screenPt = picUserIcon.PointToScreen(Point.Empty);
                int x = screenPt.X + (picUserIcon.Width - userMenu.Width) / 2;
                int y = screenPt.Y + picUserIcon.Height;
                userMenu.Show(new Point(x, y));
            }
        }

        public Func<bool> ConfirmLeaveHandler { get; set; }

        private Button CreateButton(string text, Action onClick)
        {
            var btn = new Button
            {
                Text = text,
                Width = 100,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(5),
                BackColor = Color.LightSteelBlue,
                ForeColor = Color.Black,
                Font = new Font("微軟正黑體", 9, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };

            btn.Click += (s, e) =>
            {
                if (ConfirmLeaveHandler == null || ConfirmLeaveHandler())
                    onClick();
            };

            return btn;
        }

        private void SelectTab(string tabName)
        {
            if (MainTabControl == null) return;
            foreach (TabPage tab in MainTabControl.TabPages)
            {
                if (tab.Name == tabName)
                {
                    MainTabControl.SelectedTab = tab;
                    break;
                }
            }
        }

        private void SetCircularRegion(PictureBox pic)
        {
            using (var path = new GraphicsPath())
            {
                path.AddEllipse(0, 0, pic.Width - 1, pic.Height - 1);
                pic.Region = new Region(path);
            }
        }

        public void SetLoginState(bool loggedIn, string nickname = "", string role = "User")
        {
            BtnLoginNav.Visible = !loggedIn;
            picUserIcon.Visible = loggedIn;
            currentRole = role;

            if (loggedIn)
            {
                var imgPath = Path.Combine(Application.StartupPath, "user_icon.jpg");
                picUserIcon.Image = File.Exists(imgPath)
                    ? Image.FromFile(imgPath)
                    : SystemIcons.Information.ToBitmap();
                SetCircularRegion(picUserIcon);

                LblWelcomeText.Text = $"歡迎，{nickname}";
                InitUserMenu(role);
            }
            else
            {
                picUserIcon.Image = null;
                LblWelcomeText.Text = "登入會員享有更多功能";
            }
        }
    }
}
