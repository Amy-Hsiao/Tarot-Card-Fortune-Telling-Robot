using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Media;
using System.Drawing.Imaging;
using OpenAI.Chat;
using System.Text.RegularExpressions;
using System.Windows.Forms.VisualStyles;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using DotNetEnv;

namespace 第一頁
{
    
    public partial class Form1 : Form
    {
        string apiKey;
        private NavigationBar navBar;
        private bool isLoggedIn = false;
        private string loggedInUser = "";
        private string loggedInAccount = "";
        private ToolTip cardToolTip = new ToolTip();
        private Dictionary<PictureBox, Size> originalSizes = new Dictionary<PictureBox, Size>();
        private const float scaleFactor = 2f;
        SoundPlayer killswitch = new SoundPlayer("killswitch.wav");
        SoundPlayer swimming = new SoundPlayer("swimming.wav");
        SoundPlayer Beginning = new SoundPlayer("The Beginning.wav");
        SoundPlayer Immaterial = new SoundPlayer("Immaterial.wav");
        SoundPlayer dream = new SoundPlayer("dream.wav");
        public Form1()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            
            navBar = new NavigationBar();
            navBar.MainTabControl = tabControlMain;
            navBar.ConfirmLeaveHandler = ConfirmLeaveMemberCenter;
            navBar.BtnLoginNav.Click += BtnLoginNav_Click;
            navBar.Dock = DockStyle.Top;
            this.Controls.Add(navBar);
            navBar.BringToFront();

            tabControlMain.SelectedIndexChanged += (s, e) =>
            {
                var pg = tabControlMain.SelectedTab;
                if (pg.BackgroundImage != null)
                {
                    this.BackgroundImage = pg.BackgroundImage;
                    this.BackgroundImageLayout = pg.BackgroundImageLayout;
                }
            };

            navBar.LogoutClicked += (s, e) =>
            {
                SetLoginState(false);
                ResetLoginFields();
                tabControlMain.SelectedTab = tabPageHome; // 登出後返回首頁
                MessageBox.Show("您已登出，歡迎再次登入！");
            };
            navBar.MemberCenterClicked += (s, e) =>
            {
                if (!ConfirmLeaveMemberCenter()) return;
                tabControlMain.SelectedTab = tabPageMemberCenter;
            };
            navBar.AdminPanelClicked += (s, e) =>
            {
                if (!ConfirmLeaveMemberCenter()) return;
                tabControlMain.SelectedTab = tabPageAdminPanel;
            };
            navBar.AppointmentClicked += (s, e) =>
            {
                if (!ConfirmLeaveMemberCenter()) return;
                tabControlMain.SelectedTab = tabPageAppointment;
            };
            navBar.AppointmentAdminClicked += (s, e) =>
            {
                if (!ConfirmLeaveMemberCenter()) return;
                tabControlMain.SelectedTab = tabPageAppointmentAdmin;
            };

            btnMemberEditSave.Click -= btnMemberEditSave_Click;
            btnMemberEditSave.Click += btnMemberEditSave_Click;
            tabPageMemberCenter.Enter += tabPageMemberCenter_Enter;
            tabControlMain.Selecting += TabControlMain_Selecting;
            btnResetDraw.Click -= btnResetDraw_Click;
            btnResetDraw.Click += btnResetDraw_Click;
            btnViewResult.Click -= btnViewResult_Click;
            btnViewResult.Click += btnViewResult_Click;
            btnAIReader1.Click += AIReader_Click;
            btnAIReader2.Click += AIReader_Click;
            btnAIReader3.Click += AIReader_Click;
            btnAIReader4.Click += AIReader_Click;
            btnAIReader5.Click += AIReader_Click;
            btnAIReader6.Click += AIReader_Click;
            btnAIReader4.MouseHover += LockedHover;
            btnAIReader5.MouseHover += LockedHover;
            btnAIReader6.MouseHover += LockedHover;
            btnAIReader4.MouseLeave += LockedLeave;
            btnAIReader5.MouseLeave += LockedLeave;
            btnAIReader6.MouseLeave += LockedLeave;
            tabControlMain.Selecting += (s, e) =>
            {
                _unlockTip.Hide(btnAIReader4);
                _unlockTip.Hide(btnAIReader5);
                _unlockTip.Hide(btnAIReader6);
            };
            btnSingleTarot.Click += (s, e) =>
            {
                isYesNoMode = true;
                drawCardCount = 1;
                tabControlMain.SelectedTab = tabPageAskQuestion;
            };
            btnThreeTarot.Click += (s, e) =>
            {
                isYesNoMode = true;
                drawCardCount = 3;
                tabControlMain.SelectedTab = tabPageAskQuestion;
            };
            SetupFortuneReaderRadioButtons();
            btnThreeTarot.Enabled = isLoggedIn;
            tabPageAppointment.Enter -= tabPageAppointment_Enter;
            tabPageAppointment.Enter += tabPageAppointment_Enter;
            dgvAppointmentSchedule.CellDoubleClick -= dgvAppointmentSchedule_CellDoubleClick;
            dgvAppointmentSchedule.CellDoubleClick += dgvAppointmentSchedule_CellDoubleClick;
        }
        
        /* 首頁 */
        private void Form1_Load(object sender, EventArgs e)
        {
            Env.Load();
            apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (File.Exists(@"sunset.jpg"))
            {
                tabPageHome.BackgroundImage = Image.FromFile(@"sunset.jpg");
                tabPageHome.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageLogin.BackgroundImage = Image.FromFile(@"sunset.jpg");
                tabPageLogin.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageRegister.BackgroundImage = Image.FromFile(@"sunset.jpg");
                tabPageRegister.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageAdminPanel.BackgroundImage = Image.FromFile(@"sunset.jpg");
                tabPageAdminPanel.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageMemberCenter.BackgroundImage = Image.FromFile(@"sunset.jpg");
                tabPageMemberCenter.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageAppointment.BackgroundImage = Image.FromFile(@"sunset.jpg");
                tabPageAppointment.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageAppointmentAdmin.BackgroundImage = Image.FromFile(@"sunset.jpg");
                tabPageAppointmentAdmin.BackgroundImageLayout = ImageLayout.Stretch;

            }

            if (File.Exists(@"ocean.jpg"))
            {
                tabPageAITarot.BackgroundImage = Image.FromFile(@"ocean.jpg");
                tabPageAITarot.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageYesNoTarot.BackgroundImage = Image.FromFile(@"ocean.jpg");
                tabPageYesNoTarot.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageTimeTarot.BackgroundImage = Image.FromFile(@"ocean.jpg");
                tabPageTimeTarot.BackgroundImageLayout = ImageLayout.Stretch;
                tabPageAskQuestion.BackgroundImage = Image.FromFile(@"ocean.jpg");
                tabPageAskQuestion.BackgroundImageLayout = ImageLayout.Stretch;
            }

            this.BackgroundImage = tabPageHome.BackgroundImage;
            this.BackgroundImageLayout = tabPageHome.BackgroundImageLayout;

            if (!File.Exists("users.csv"))
            {
                var header = "暱稱,帳號,密碼,性別,生日,Email,角色,狀態" + Environment.NewLine;
                File.WriteAllText("users.csv", header, new UTF8Encoding(true));
            }

            if (File.Exists(@"question.jpg"))
            {
                tabPageAskQuestion.BackgroundImage = Image.FromFile(@"question.jpg");
            }

            if (File.Exists(@"background.jpg"))
            {
                tabPageFortuneDraw.BackgroundImage = Image.FromFile(@"background.jpg");
            }

            if (File.Exists(@"sunrise.jpg"))
            {
                tabPageResult.BackgroundImage = Image.FromFile(@"sunrise.jpg");
                tabPageYesNoResult.BackgroundImage = Image.FromFile(@"sunrise.jpg");
                tabPageResult.BackgroundImage = Image.FromFile(@"sunrise.jpg");
                tabPageFortuneResult.BackgroundImage = Image.FromFile(@"sunrise.jpg");
            }

            swimming.Play();

            tabControlMain.Appearance = TabAppearance.FlatButtons;
            tabControlMain.ItemSize = new Size(0, 1);
            tabControlMain.SizeMode = TabSizeMode.Fixed;

            txtLoginPassword.UseSystemPasswordChar = true;
            txtRegisterPassword.UseSystemPasswordChar = true;
            checkboxShowLoginPwd.CheckedChanged += (s, eArgs) =>
            {
                txtLoginPassword.UseSystemPasswordChar = !checkboxShowLoginPwd.Checked;
            };
            checkboxShowRegisterPwd.CheckedChanged += (s, eArgs) =>
            {
                txtRegisterPassword.UseSystemPasswordChar = !checkboxShowRegisterPwd.Checked;
            };

            label10.BackColor = Color.Transparent;
            label11.BackColor = Color.LightCyan;
            label12.BackColor = Color.LightCyan;
            label13.BackColor = Color.AliceBlue;
            label14.BackColor = Color.AliceBlue;
            label15.BackColor = Color.LavenderBlush;
            label16.BackColor = Color.LavenderBlush;
        }

        // AI塔羅占卜button
        private void btnAI_Click(object sender, EventArgs e)
        {
            killswitch.Play();
            if (!ConfirmLeaveMemberCenter()) return;
            tabControlMain.SelectedTab = tabPageAITarot;
        }

        // 是否塔羅占卜button
        private void buttonYesNo_Click(object sender, EventArgs e)
        {
            Beginning.Play();
            if (!ConfirmLeaveMemberCenter()) return;
            tabControlMain.SelectedTab = tabPageYesNoTarot;
        }
        // 塔羅運勢button
        private void buttonTime_Click(object sender, EventArgs e)
        {
            Immaterial.Play();
            if (!ConfirmLeaveMemberCenter()) return;
            tabControlMain.SelectedTab = tabPageTimeTarot;
        }

        /* 登入頁面 */
        private void tabPageLogin_Enter(object sender, EventArgs e)
        {
            label42.BackColor = Color.Transparent;
            label1.BackColor = Color.Transparent;
            label2.BackColor = Color.Transparent;
            label34.BackColor = Color.Transparent;
            label9.BackColor = Color.Transparent;
            label28.BackColor = Color.Transparent;
            label61.BackColor = Color.Transparent;
            checkboxShowLoginPwd.BackColor = Color.Transparent;
            radioButton1.BackColor = Color.Transparent;
        }
        private void btnLoginOrRegister_Click(object sender, EventArgs e)
        {
            
            string acc = txtLoginAccount.Text.Trim();
            string pwd = txtLoginPassword.Text.Trim();

            var result = TryLogin(acc, pwd);
            if (result == LoginResult.AccountNotFound)
            {
                // 帳號不存在 → 前往註冊頁
                if (!ConfirmLeaveMemberCenter()) return;
                txtRegisterAccount.Text = acc;
                txtRegisterPassword.Text = pwd;
                txtRegisterAccount.ReadOnly = true;
                txtRegisterPassword.ReadOnly = true;
                tabControlMain.SelectedTab = tabPageRegister;
            }
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            label61.Visible = radioButton1.Checked;
            textBox4.Visible = radioButton1.Checked;
            button4.Visible = radioButton1.Checked;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            string toEmail = textBox4.Text.Trim();

            if (string.IsNullOrEmpty(toEmail) || !toEmail.Contains("@"))
            {
                MessageBox.Show("請輸入正確的 Gmail！");
                return;
            }
            if (!File.Exists("users.csv"))
            {
                MessageBox.Show("使用者資料不存在！");
                return;
            }
            var lines = File.ReadAllLines("users.csv", Encoding.UTF8).Skip(1);
            var found = lines.FirstOrDefault(line => line.Split(',')[5] == toEmail);

            if (found == null)
            {
                MessageBox.Show("找不到此註冊信箱！");
                return;
            }

            string password = found.Split(',')[2]; // 第三欄是密碼
            SendForgetPasswordLetter(toEmail,password);
        }
        public void SendForgetPasswordLetter(string toEmail,string password)
        {
            string from = "liyingxiao992@gmail.com";
            string appPassword = "uqzt piqb balp gqpq";

            string subject = "【塔羅占卜】密碼驗證信";

            string htmlBody = $@"
            <html><body style='font-family:微軟正黑體'>
            <h2 style='color:#8A2BE2;'>🔐 您的帳號密碼資訊</h2>
            <p>您在塔羅占卜平台註冊的密碼為：</p>
            <p style='font-size:18px; color:#d63384;'><strong>{password}</strong></p>
            <hr />
            <p style='color:#888;'>🌟 願宇宙給你溫柔的指引 ✨</p>
            <p>— 你的塔羅占卜小幫手 💜</p>
            </body></html>";

            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(from, "塔羅占卜機器人");
            mail.To.Add(toEmail);
            mail.Subject = subject;
            mail.Body = htmlBody;
            mail.IsBodyHtml = true;

            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            smtp.Credentials = new NetworkCredential(from, appPassword);
            smtp.EnableSsl = true;

            try
            {
                smtp.Send(mail);
                MessageBox.Show("💌 寄出成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("寄信失敗：" + ex.Message);
            }
        }
        // 登出
        private void BtnLoginNav_Click(object sender, EventArgs e)
        {
            if (isLoggedIn)
            {
                // 登出不需要檢查離開
                SetLoginState(false);
                ResetLoginFields();
                tabControlMain.SelectedTab = tabPageHome;
                MessageBox.Show("您已登出，歡迎再次登入！");
            }
            else
            {
                // 要跳去登入頁，也要先確認
                if (!ConfirmLeaveMemberCenter()) return;
                ResetLoginFields();
                tabControlMain.SelectedTab = tabPageLogin;
            }
        }

        enum LoginResult
        {
            Success,
            AccountNotFound,
            PasswordIncorrect,
            Disabled
        }

        private LoginResult TryLogin(string acc, string pwd)
        {
            if (string.IsNullOrEmpty(acc) || string.IsNullOrEmpty(pwd))
            {
                MessageBox.Show("帳號與密碼皆不得為空白！", "登入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return LoginResult.PasswordIncorrect;
            }

            if (!Regex.IsMatch(acc, @"^[a-zA-Z0-9\p{P}]{1,20}$") || !Regex.IsMatch(pwd, @"^[a-zA-Z0-9\p{P}]{1,20}$"))
            {
                MessageBox.Show("帳號與密碼僅限 1~20 字元，英數字與符號組合");
                return LoginResult.PasswordIncorrect;
            }

            if (!File.Exists("users.csv"))
            {
                MessageBox.Show("使用者資料檔不存在。請先註冊！");
                return LoginResult.AccountNotFound;
            }

            var lines = File.ReadAllLines("users.csv", Encoding.UTF8).Skip(1);
            var found = lines.FirstOrDefault(line => line.Split(',')[1] == acc);

            if (found == null)
            {
                return LoginResult.AccountNotFound;
            }

            var parts = found.Split(',');
            string storedPassword = parts[2];
            string role = parts[6];
            string status = parts[7];

            if (pwd != storedPassword)
            {
                MessageBox.Show("此帳號已被註冊且密碼錯誤！");
                return LoginResult.PasswordIncorrect;
            }

            if (status == "Disabled")
            {
                MessageBox.Show("此帳號已被停權，如有疑問請聯繫管理者，謝謝！");
                tabControlMain.SelectedTab = tabPageHome;
                return LoginResult.Disabled;
            }

            // 登入成功
            isLoggedIn = true;
            loggedInUser = parts[0];
            loggedInAccount = parts[1];
            currentUserAccount = acc;
            navBar.SetLoginState(true, loggedInUser, role);
            MessageBox.Show($"您好 {loggedInUser}，登入成功！");

            if (role == "Admin")
            {
                tabControlMain.SelectedTab = tabPageAdminPanel;
                LoadUsersToGrid();
                FormatUserGrid();
            }
            else
            {
                tabControlMain.SelectedTab = tabPageHome;
            }

            return LoginResult.Success;
        }

        private void SetLoginState(bool isLogin)
        {
            isLoggedIn = isLogin;
            navBar.SetLoginState(isLogin, loggedInUser);
            
            // 如果正在塔羅運勢頁面 更新塔羅師radiobuttom
            if (tabControlMain.SelectedTab == tabPageTimeTarot)
            {
                InitializeTarotReaderStyles();
            }
        }

        private void ResetLoginFields()
        {
            txtLoginAccount.Clear();
            txtLoginPassword.Clear();
            txtRegisterAccount.Clear();
            txtRegisterPassword.Clear();
            txtNickname.Clear();
            txtEmail.Clear();
            rdoMale.Checked = false;
            rdoFemale.Checked = false;
            checkboxShowLoginPwd.Checked = false;
            checkboxShowRegisterPwd.Checked = false;
        }

        /* 註冊頁面 */
        private void tabPageRegister_Enter(object sender, EventArgs e)
        {
            label41.BackColor = Color.Transparent;
            label3.BackColor = Color.Transparent;
            label4.BackColor = Color.Transparent;
            label5.BackColor = Color.Transparent;
            label6.BackColor = Color.Transparent;
            label7.BackColor = Color.Transparent;
            label8.BackColor = Color.Transparent;
            checkboxShowRegisterPwd.BackColor = Color.Transparent;
            rdoMale.BackColor = Color.Transparent;
            rdoFemale.BackColor = Color.Transparent;
        }
        private void btnRegister_Click(object sender, EventArgs e)
        {
            string acc = txtRegisterAccount.Text.Trim();
            string pwd = txtRegisterPassword.Text.Trim();
            string nickname = txtNickname.Text.Trim();
            string email = txtEmail.Text.Trim();
            string gender = rdoMale.Checked ? "生理男" : (rdoFemale.Checked ? "生理女" : "");
            string birthday = dtpBirthday.Value.ToString("yyyy/MM/dd");
            string role = "User";
            string status = "Active";

            var lines = File.Exists("users.csv")
            ? File.ReadAllLines("users.csv", Encoding.UTF8).Skip(1).ToList()
            : new List<string>();

            // 驗證：暱稱
            if (string.IsNullOrEmpty(nickname) || nickname.Length > 10)
            {
                MessageBox.Show("暱稱必填，且長度不得超過10字元");
                return;
            }
            if (lines.Any(line => line.Split(',')[0] == nickname))
            {
                MessageBox.Show($"暱稱「{nickname}」已被使用，請更換其他暱稱！");
                return;
            }

            // 驗證：Email
            if (!Regex.IsMatch(email, @"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"))
            {
                MessageBox.Show("Email 格式不正確，請輸入正確的 Email 格式");
                return;
            }

            // 驗證：性別
            if (string.IsNullOrEmpty(gender))
            {
                MessageBox.Show("請選擇性別！");
                return;
            }

            // 驗證：帳號是否已存在
            if (lines.Any(line => line.Split(',')[1] == acc))
            {
                MessageBox.Show("此帳號已存在！");
                return;
            }

            // 儲存至 CSV（格式：暱稱,帳號,密碼,性別,生日,email,角色,狀態）
            string newLine = $"{nickname},{acc},{pwd},{gender},{birthday},{email},{role},{status}" + Environment.NewLine;
            File.AppendAllText("users.csv", newLine, new System.Text.UTF8Encoding(true));

            // 登入狀態設定
            if (!ConfirmLeaveMemberCenter()) return;
            isLoggedIn = true;
            loggedInUser = nickname;
            loggedInAccount = acc;
            currentUserAccount = acc;
            navBar.SetLoginState(true, loggedInUser);
            MessageBox.Show($"您好 {loggedInUser}，註冊成功並已登入！");
            tabControlMain.SelectedTab = tabPageHome;
            SendRegisterEmail(email,nickname);
        }
        public void SendRegisterEmail(string email,string nickname)
        {
            string from = "liyingxiao992@gmail.com";
            string appPassword = "uqzt piqb balp gqpq";
            string subject = "🔮 歡迎加入塔羅占卜的魔法旅程 ✨";

            string htmlBody = $@"
            <html>
            <body style='font-family:微軟正黑體; line-height:1.6;'>
            <h2 style='color:#8A2BE2;'>✨ 親愛的 {nickname}，歡迎來到塔羅占卜世界！</h2>
            <p>感謝您註冊塔羅占卜平台！</p>
            <p>在這裡，您可以透過占卜牌卡探索自我、聆聽內心的聲音，找到屬於自己的指引。</p>
            <p>無論是想了解感情、工作、學業，還是生活的每個十字路口，我們都會在這裡陪伴您。</p>

            <div style='padding:12px; border:1px solid #8A2BE2; border-radius:8px; background-color:#f8f4ff;'>
            <p style='margin:0;'><strong>🔮 接下來您可以：</strong></p>
            <ul>
                <li>點選 <strong>開始占卜</strong> → 提出問題 → 抽牌</li>
                <li>選擇 <strong>AI塔羅 or 是/否占卜 or 塔羅運勢</strong></li>
                <li>體驗療癒系塔羅解讀，為你帶來指引與祝福</li>
            </ul>
        </div>
            <p style='color:#888;'>🌟 願宇宙給你溫柔的指引 ✨</p>
            <p>— 你的塔羅占卜小幫手 💜</p>
            </body></html>";

            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(from, "塔羅占卜機器人");
            mail.To.Add(email);
            mail.Subject = subject;
            mail.Body = htmlBody;
            mail.IsBodyHtml = true;

            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            smtp.Credentials = new NetworkCredential(from, appPassword);
            smtp.EnableSsl = true;

            try
            {
                smtp.Send(mail);
                //MessageBox.Show("💌 已寄歡迎信件給您囉！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("寄信失敗：" + ex.Message);
            }
        }

        /* 管理者後台 */
        private bool isDataModified = false;
        private int previousComboIndex = 0;

        private void tabPageAdminPanel_Enter(object sender, EventArgs e)
        {
            label57.BackColor = Color.Transparent;
            label56.BackColor = Color.Transparent;
            label29.BackColor = Color.Transparent;
            if (comboFilterRole.Items.Count == 0)
            {
                comboFilterRole.Items.AddRange(new string[] { "全部", "使用者", "管理者" });
                comboFilterRole.SelectedIndex = 0;
                previousComboIndex = 0;
                comboFilterRole.SelectedIndexChanged += comboFilterRole_SelectedIndexChanged;
            }
            LoadUsersToGrid("全部");
            FormatUserGrid();
            isDataModified = false;
        }
        
        private void comboFilterRole_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isDataModified)
            {
                var result = MessageBox.Show("尚未儲存的修改會遺失，確定切換？", "提醒", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    // 取消切換 → 還原選項
                    comboFilterRole.SelectedIndexChanged -= comboFilterRole_SelectedIndexChanged;
                    comboFilterRole.SelectedIndex = previousComboIndex; // 需要記住上一次選項
                    comboFilterRole.SelectedIndexChanged += comboFilterRole_SelectedIndexChanged;
                    return;
                }
            }
            // 如果使用者確定切換 → 載入新的角色列表
            string selected = comboFilterRole.SelectedItem.ToString();
            LoadUsersToGrid(selected);
            FormatUserGrid();
            isDataModified = false;
            previousComboIndex = comboFilterRole.SelectedIndex;
        }
        
        private void dataGridViewUsers_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            isDataModified = true;
        }

        private void LoadUsersToGrid(string filterRole = "全部")
        {
            if (!File.Exists("users.csv")) return;

            var lines = File.ReadAllLines("users.csv", Encoding.UTF8).ToList();
            if (lines.Count == 0) return;

            DataTable dt = new DataTable();
            string[] headers = lines[0].Split(',');

            foreach (var header in headers)
                dt.Columns.Add(header);

            foreach (var line in lines.Skip(1))
            {
                var parts = line.Split(',');
                if (filterRole == "全部" ||
                    (filterRole == "使用者" && parts[6] == "User") ||
                    (filterRole == "管理者" && parts[6] == "Admin"))
                {
                    // 格式化生日為 yyyy/MM/dd
                    if (DateTime.TryParse(parts[4], out DateTime parsedDate))
                    {
                        parts[4] = parsedDate.ToString("yyyy/MM/dd");
                    }
                    dt.Rows.Add(parts);
                }
            }
            dataGridViewUsers.AllowUserToAddRows = (filterRole == "管理者");
            dataGridViewUsers.DataSource = dt;
            dataGridViewUsers.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridViewUsers.RowHeadersVisible = false;
        }

        private void FormatUserGrid()
        {
            if (dataGridViewUsers.Columns.Count == 0) return;

            foreach (DataGridViewColumn col in dataGridViewUsers.Columns)
            {
                if (col.HeaderText == "角色" || col.HeaderText == "狀態")
                    col.ReadOnly = false;
                else
                    col.ReadOnly = !dataGridViewUsers.AllowUserToAddRows;  // 管理者時可編輯其餘欄位
            }

            // 設定角色與狀態欄為下拉選單
            if (dataGridViewUsers.Columns.Contains("角色"))
            {
                var roleColumn = new DataGridViewComboBoxColumn
                {
                    DataPropertyName = "角色",
                    HeaderText = "角色",
                    Name = "角色",
                    DataSource = new string[] { "User", "Admin" },
                    FlatStyle = FlatStyle.Flat
                };
                int roleIndex = dataGridViewUsers.Columns["角色"].Index;
                dataGridViewUsers.Columns.RemoveAt(roleIndex);
                dataGridViewUsers.Columns.Insert(roleIndex, roleColumn);
            }

            if (dataGridViewUsers.Columns.Contains("狀態"))
            {
                var statusColumn = new DataGridViewComboBoxColumn
                {
                    DataPropertyName = "狀態",
                    HeaderText = "狀態",
                    Name = "狀態",
                    DataSource = new string[] { "Active", "Disabled" },
                    FlatStyle = FlatStyle.Flat
                };
                int statusIndex = dataGridViewUsers.Columns["狀態"].Index;
                dataGridViewUsers.Columns.RemoveAt(statusIndex);
                dataGridViewUsers.Columns.Insert(statusIndex, statusColumn);
            }
            dataGridViewUsers.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            foreach (DataGridViewColumn col in dataGridViewUsers.Columns)
            {
                if (col.HeaderText == "Email")
                {
                    col.FillWeight = 200; // Email 欄寬一點
                }
                else
                {
                    col.FillWeight = 100; // 其他欄維持正常
                }
            }
        }

        private void btnSaveChanges_Click(object sender, EventArgs e)
        {
            string roleSelected = comboFilterRole.SelectedItem?.ToString() ?? "全部";
            string defaultFileName = "users.csv";
            if (roleSelected == "使用者") defaultFileName = "User.csv";
            else if (roleSelected == "管理者") defaultFileName = "Admin.csv";

            var sb = new StringBuilder();

            var allLines = File.Exists("users.csv") ? File.ReadAllLines("users.csv", Encoding.UTF8).ToList() : new List<string>();
            if (allLines.Count == 0) return;
            var headerLine = allLines[0];
            var existingData = allLines.Skip(1).ToList();

            var updatedData = new List<string>();
            var existingAccs = new HashSet<string>();
            var existingAdminNicknames = new HashSet<string>();
            var existingAllNicknames = new HashSet<string>();

            foreach (var line in existingData)
            {
                var parts = line.Split(',');
                if (parts.Length >= 7)
                {
                    existingAccs.Add(parts[1]);
                    existingAllNicknames.Add(parts[0]);
                    if (parts[6] == "Admin") existingAdminNicknames.Add(parts[0]);
                }
            }

            var accSet = new HashSet<string>();
            var adminNickSet = new HashSet<string>();
            var allNickSet = new HashSet<string>();

            foreach (DataGridViewRow row in dataGridViewUsers.Rows)
            {
                if (row.IsNewRow) continue;

                string nickname = row.Cells["暱稱"].Value?.ToString().Trim() ?? "";
                string acc = row.Cells["帳號"].Value?.ToString().Trim() ?? "";
                string pwd = row.Cells["密碼"].Value?.ToString().Trim() ?? "";
                string gender = row.Cells["性別"].Value?.ToString().Trim() ?? "";
                string birthdayRaw = row.Cells["生日"].Value?.ToString().Trim() ?? "";
                string birthday = DateTime.TryParse(birthdayRaw, out DateTime bday) ? bday.ToString("yyyy/MM/dd") : birthdayRaw;
                string email = row.Cells["Email"].Value?.ToString().Trim() ?? "";
                string role = row.Cells["角色"].Value?.ToString().Trim() ?? "";
                string status = row.Cells["狀態"].Value?.ToString().Trim() ?? "";

                if (string.IsNullOrEmpty(nickname) || nickname.Length > 10)
                {
                    MessageBox.Show($"暱稱必填，且長度不得超過10字元");
                    return;
                }
                if (!allNickSet.Add(nickname))
                {
                    MessageBox.Show($"暱稱「{nickname}」已被使用！");
                    return;
                }
                if (existingAllNicknames.Contains(nickname) && !existingData.Any(line => line.Split(',')[0] == nickname && line.Split(',')[1] == acc))
                {
                    MessageBox.Show($"暱稱「{nickname}」已被使用！");
                    return;
                }

                if (!Regex.IsMatch(acc, @"^[a-zA-Z0-9\p{P}]{1,20}$"))
                {
                    MessageBox.Show("帳號必填，且僅限 1~20 字元，英數字與符號組合");
                    return;
                }
                if (!accSet.Add(acc))
                {
                    MessageBox.Show($"帳號「{acc}」已被註冊過！");
                    return;
                }
                if (existingAccs.Contains(acc) && !existingData.Any(line => line.Split(',')[1] == acc && line.Split(',')[0] == nickname))
                {
                    MessageBox.Show($"帳號「{acc}」已被註冊過！");
                    return;
                }

                if (!Regex.IsMatch(pwd, @"^[a-zA-Z0-9\p{P}]{1,20}$"))
                {
                    MessageBox.Show("密碼必填，且僅限 1~20 字元，英數字與符號組合");
                    return;
                }

                if (gender != "生理男" && gender != "生理女")
                {
                    MessageBox.Show($"性別格式錯誤，格式為「生理男」或「生理女」");
                    return;
                }

                if (string.IsNullOrEmpty(birthday) || !DateTime.TryParseExact(birthday, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out _))
                {
                    MessageBox.Show($"生日格式錯誤，格式為「YYYY/MM/DD」");
                    return;
                }

                if (!Regex.IsMatch(email, @"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"))
                {
                    MessageBox.Show("Email 格式不正確，請輸入正確的 Email 格式");
                    return;
                }

                if (role != "User" && role != "Admin")
                {
                    MessageBox.Show($"請選擇有效的角色（User 或 Admin）");
                    return;
                }
                if (status != "Active" && status != "Disabled")
                {
                    MessageBox.Show($"請選擇有效的狀態（Active 或 Disabled）");
                    return;
                }

                updatedData.Add($"{nickname},{acc},{pwd},{gender},{birthday},{email},{role},{status}");
            }

            // 組合成新的 users.csv 資料
            var fullSb = new StringBuilder();
            fullSb.AppendLine(headerLine);
            var mergedData = existingData
                .Where(line => !updatedData.Any(u => line.Split(',')[1] == u.Split(',')[1]))
                .ToList();
            mergedData.AddRange(updatedData);
            foreach (var line in mergedData)
            {
                fullSb.AppendLine(line);
            }

            // 匯出資料
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Title = "請選擇匯出位置",
                FileName = defaultFileName,
                Filter = "CSV 檔案 (*.csv)|*.csv",
                InitialDirectory = Application.StartupPath
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                // 寫入 users.csv
                File.WriteAllText("users.csv", fullSb.ToString(), new UTF8Encoding(true));

                // 寫入匯出的檔案（若為 Admin.csv 或 User.csv 則只匯出該角色）
                if (defaultFileName == "Admin.csv")
                {
                    var adminOnly = fullSb.ToString().Split('\n')
                        .Where(line => line.Contains(",Admin"))
                        .Prepend(headerLine);
                    File.WriteAllText(saveDialog.FileName, string.Join("\n", adminOnly), new UTF8Encoding(true));
                }
                else if (defaultFileName == "User.csv")
                {
                    var userOnly = fullSb.ToString().Split('\n')
                        .Where(line => line.Contains(",User"))
                        .Prepend(headerLine);
                    File.WriteAllText(saveDialog.FileName, string.Join("\n", userOnly), new UTF8Encoding(true));
                }
                else
                {
                    File.WriteAllText(saveDialog.FileName, fullSb.ToString(), new UTF8Encoding(true));
                }

                MessageBox.Show("帳號資料已成功儲存與匯出！");
                isDataModified = false;
            }
        }

        /* AI塔羅占卜頁面 */
        string selectedAIReader = "";
        private Dictionary<string, string> readerStyles = new Dictionary<string, string>();
        private string style = "";
        private string userQuestion = "";
        private readonly ToolTip _unlockTip = new ToolTip
        {
            ShowAlways = true,
            InitialDelay = 500,
            ReshowDelay = 100,
            AutoPopDelay = 3000
        };

        private void tabPageAITarot_Enter(object sender, EventArgs e)
        {
            label25.BackColor = Color.LightCyan;
            label17.BackColor = Color.LightCyan;
            label22.BackColor = Color.Transparent;
            label23.BackColor = Color.Transparent;
            label24.BackColor = Color.Transparent;
            label31.BackColor = Color.Transparent;
            label46.BackColor = Color.Transparent;
            label47.BackColor = Color.Transparent;
            label18.BackColor = Color.Transparent;
            label19.BackColor = Color.Transparent;
            label20.BackColor = Color.Transparent;
            label30.BackColor = Color.Transparent;
            label44.BackColor = Color.Transparent;
            label45.BackColor = Color.Transparent;
            label49.BackColor = Color.Transparent;
            label50.BackColor = Color.Transparent;
            label51.BackColor = Color.Transparent;

            var exeDir = Application.StartupPath;
            var thumbSize = new Size(80, 100);
            var imgMap = new Dictionary<Button, string>
            {
                [btnAIReader1] = "readerA.png",  // A 療癒系
                [btnAIReader2] = "readerB.png",  // B 現實系
                [btnAIReader3] = "readerC.png",  // C 神秘系
                [btnAIReader4] = "readerD.png",  // D 心靈共鳴
                [btnAIReader5] = "readerE.png",  // E 未來趨勢
                [btnAIReader6] = "readerF.png",  // F 可愛治療
            };

            var keyMap = new Dictionary<Button, string>
            {
                [btnAIReader1] = "reader1",
                [btnAIReader2] = "reader2",
                [btnAIReader3] = "reader3",
                [btnAIReader4] = "reader4",
                [btnAIReader5] = "reader5",
                [btnAIReader6] = "reader6",
            };

            foreach (var kv in imgMap)
            {
                var btn = kv.Key;
                var file = Path.Combine(exeDir, kv.Value);
                if (File.Exists(file))
                {
                    btn.Text = "";
                    btn.Size = thumbSize;
                    btn.BackgroundImageLayout = ImageLayout.Zoom;
                    btn.BackgroundImage = Image.FromFile(file);
                }
                // 一定要設定 Tag，AIReader_Click 才能讀到正確的 key
                btn.Tag = keyMap[btn];
            }

            if (readerStyles.Count == 0)
            {
                // 免費
                readerStyles["reader1"] = "療癒系塔羅占卜師，擅長用溫柔的語氣解釋牌義，會用比喻、鼓勵與祝福的方式說明";
                readerStyles["reader2"] = "現實系塔羅占卜師，重視直觀與邏輯，解牌時會坦率地指出問題核心，提供具體可行的建議";
                readerStyles["reader3"] = "神秘系塔羅占卜師，注重牌與宇宙能量的連結，擅長從靈性視角詮釋命運的流動與內在課題";
                // 會員限定
                readerStyles["reader4"] = "心靈共鳴系塔羅占卜師，專注傾聽內心深層情感，透過塔羅釋放壓抑情緒，帶來情感療癒與自我成長";
                readerStyles["reader5"] = "未來趨勢系塔羅占卜師，結合占星能量與塔羅牌陣，剖析未來走向與關鍵時機，幫助掌握人生轉折與契機";
                readerStyles["reader6"] = "可愛治療系塔羅占卜師，擅長以柔和可愛的卡通符號與色彩，結合塔羅牌意象，輕鬆撫慰心靈";
            }

            foreach (var btn in new[] { btnAIReader1, btnAIReader2, btnAIReader3 })
            {
                btn.BackColor = Color.LightSteelBlue;
                btn.Cursor = Cursors.Hand;
            }

            foreach (var btn in new[] { btnAIReader4, btnAIReader5, btnAIReader6 })
            {
                if (!isLoggedIn)
                {
                    btn.BackColor = Color.LightGray;
                    btn.Cursor = Cursors.No;
                }
                else
                {
                    btn.BackColor = Color.LightSteelBlue;
                    btn.Cursor = Cursors.Hand;
                }
            }
        }
        private void LockedHover(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if (btn != null && !isLoggedIn)
            {
                _unlockTip.Show(
                    "請先登入會員以解鎖此塔羅師",
                    btn,
                    btn.Width / 2,
                    btn.Height + 2
                );
            }
        }

        private void LockedLeave(object sender, EventArgs e)
        {
            if (sender is Control c)
                _unlockTip.Hide(c);
        }

        private void AIReader_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;

            var key = btn.Tag as string;
            if (!isLoggedIn && (key == "reader4" || key == "reader5" || key == "reader6"))
            {
                _unlockTip.Show(
                    "請先登入會員以解鎖此塔羅師",
                    btn,
                    btn.Width / 2,
                    btn.Height + 2
                );
                return;
            }

            if (readerStyles.TryGetValue(key, out var s))
            {
                style = s;
                selectedAIReader = key;
                tabControlMain.SelectedTab = tabPageAskQuestion;
            }
            else
            {
                MessageBox.Show("找不到對應的塔羅師風格設定！");
            }
        }

        /* 是否塔羅占卜 */
        private int RequiredCardCount = 3;  // 預設三張
        private void tabPageYesNoTarot_Enter(object sender, EventArgs e)
        {
            label26.BackColor = Color.AliceBlue;
            label52.BackColor = Color.AliceBlue;
            label68.BackColor = Color.Transparent;
        }
        private void btnSingleTarot_Click(object sender, EventArgs e)
        {
            
            RequiredCardCount = 1;
            tabControlMain.SelectedTab = tabPageAskQuestion;
        }
        private void btnThreeTarot_Click(object sender, EventArgs e)
        {
            
            RequiredCardCount = 3;
            tabControlMain.SelectedTab = tabPageAskQuestion;
        }

        /* 塔羅牌問題頁面 */
        private void tabPageAskQuestion_Enter(object sender, EventArgs e)
        {
            txtUserQuestion.Clear();
            label21.BackColor = Color.Thistle;
            lblCharCount.BackColor = Color.Thistle;
        }

        private void txtUserQuestion_TextChanged(object sender, EventArgs e)
        {
            if (txtUserQuestion.Text.Length > 200)
            {
                txtUserQuestion.Text = txtUserQuestion.Text.Substring(0, 200);
                txtUserQuestion.SelectionStart = txtUserQuestion.Text.Length;
                MessageBox.Show("最多只能輸入200字！");
            }
        }

        // 送出button
        private void btnSubmitQuestion_Click(object sender, EventArgs e)
        {
            userQuestion = txtUserQuestion.Text.Trim();
            if (string.IsNullOrWhiteSpace(userQuestion))
            {
                MessageBox.Show("請輸入想詢問的問題（最多200字）！");
                return;
            }
            RequiredCardCount = drawCardCount;
            ResetDraw();
            lblShow.BackColor = Color.Thistle;
            lblShow.Text = $"請選擇 {RequiredCardCount} 張牌";
            tabControlMain.SelectedTab = tabPageDrawCard;
        }

        // 提示剩餘可輸入字數
        private void txtUserQuestion_TextChanged_1(object sender, EventArgs e)
        {
            int maxLength = 200;
            int remaining = maxLength - txtUserQuestion.Text.Length;
            if (remaining < 0) remaining = 0;

            lblCharCount.Text = $"還可輸入 {remaining} 字";

            // 若超出範圍，強制裁剪文字
            if (txtUserQuestion.Text.Length > maxLength)
            {
                txtUserQuestion.Text = txtUserQuestion.Text.Substring(0, maxLength);
                txtUserQuestion.SelectionStart = txtUserQuestion.Text.Length;
            }
        }

        /* 抽牌頁面 */
        private Dictionary<string, string> tarotCardNames = new Dictionary<string, string>
        {
            { "card_1.jpg", "愚人" },
            { "card_2.jpg", "魔術師" },
            { "card_3.jpg", "女祭司" },
            { "card_4.jpg", "皇后" },
            { "card_5.jpg", "皇帝" },
            { "card_6.jpg", "教宗" },
            { "card_7.jpg", "戀人" },
            { "card_8.jpg", "戰車" },
            { "card_9.jpg", "力量" },
            { "card_10.jpg", "隱士" },
            { "card_11.jpg", "命運之輪" },
            { "card_12.jpg", "正義" },
            { "card_13.jpg", "吊人" },
            { "card_14.jpg", "死神" },
            { "card_15.jpg", "節制" },
            { "card_16.jpg", "惡魔" },
            { "card_17.jpg", "塔" },
            { "card_18.jpg", "星星" },
            { "card_19.jpg", "月亮" },
            { "card_20.jpg", "太陽" },
            { "card_21.jpg", "審判" },
            { "card_22.jpg", "世界" }
        };
        private string cardName1 = "";
        private string cardName2 = "";
        private string cardName3 = "";
        private List<PictureBox> cardPics = new List<PictureBox>();
        private Dictionary<PictureBox, Point> originalPositions = new Dictionary<PictureBox, Point>();
        private bool _drawInitialized = false;
        private List<string> deck = new List<string>();
        private List<PictureBox> selected = new List<PictureBox>();
        private bool isYesNoMode = false;                // 是否為「是否塔羅」流程
        private int drawCardCount = 3;                   // 實際抽牌數（1 或 3）
        private List<PictureBox> YesNoselected;          // 「是否塔羅」流程所選的卡片
        private string cardYesNoName1, cardYesNoName2, cardYesNoName3;
        private void tabPageDrawCard_Enter(object sender, EventArgs e)
        {
            // 設定背景圖
            if (File.Exists("background.jpg"))
                tabPageDrawCard.BackgroundImage = Image.FromFile("background.jpg");
                tabPageDrawCard.BackgroundImageLayout = ImageLayout.Stretch;

            // 建立 22 張牌
            if (!_drawInitialized)
            {
                cardPics = new List<PictureBox>
                {
                    picCard1, picCard2, picCard3, picCard4, picCard5,
                    picCard6, picCard7, picCard8, picCard9, picCard10,
                    picCard11, picCard12, picCard13, picCard14, picCard15,
                    picCard16, picCard17, picCard18, picCard19, picCard20,
                    picCard21, pictureBox22
                };

                foreach (var pic in cardPics)
                {
                    pic.SizeMode = PictureBoxSizeMode.StretchImage;
                    pic.Click -= Pic_Click;
                    pic.Click += Pic_Click;
                    originalPositions[pic] = pic.Location;

                    if (File.Exists("card_cover.jpg"))
                        pic.Image = Image.FromFile("card_cover.jpg");

                    pic.MouseEnter += Pic_MouseEnter;
                    pic.MouseLeave += Pic_MouseLeave;
                }

                lblShow.BackColor = Color.Thistle;
                lblShow.Text = $"請抽 {RequiredCardCount} 張牌";
                btnViewResult.Enabled = false;
                btnResetDraw.Enabled = false;

                _drawInitialized = true;
            }
        }

        private void ResetDraw()
        {
            selected.Clear();
            deck = Enumerable.Range(1, 22).Select(i => $"card_{i}.jpg").ToList();
            Shuffle(deck);
            lblShow.BackColor = Color.Thistle;
            lblShow.Text = $"請抽 {RequiredCardCount} 張牌";
            btnViewResult.Enabled = false;
            btnResetDraw.Enabled = false;

            foreach (var pic in cardPics)
            {
                if (File.Exists("card_cover.jpg"))
                    pic.Image = Image.FromFile("card_cover.jpg");
                pic.Enabled = true;
                pic.Tag = null;
            }
        }

        // 翻牌特效
        private void FlipCardWithAnimation(PictureBox pic, string faceImagePath)
        {
            Timer timer = new Timer();
            int step = 0;
            int totalSteps = 10;
            int originalWidth = pic.Width;
            Point basePos = originalPositions[pic]; // 取得原始位置

            timer.Interval = 20;
            timer.Tick += (s, e) =>
            {
                step++;

                if (step <= totalSteps / 2)
                {
                    pic.Width -= originalWidth / totalSteps;
                    pic.Left = basePos.X + (originalWidth - pic.Width) / 2; // 中心對齊
                }
                else if (step == (totalSteps / 2) + 1)
                {
                    if (File.Exists(faceImagePath))
                        pic.Image = Image.FromFile(faceImagePath);
                }
                else if (step <= totalSteps)
                {
                    pic.Width += originalWidth / totalSteps;
                    pic.Left = basePos.X + (originalWidth - pic.Width) / 2; // 中心對齊
                }
                else
                {
                    // 重設為初始位置與寬度
                    pic.Width = originalWidth;
                    pic.Location = basePos;
                    timer.Stop();
                    timer.Dispose();
                }
            };
            timer.Start();
        }
        private void Pic_Click(object sender, EventArgs e)
        {
            if (selected.Count >= RequiredCardCount)
                return;

            PictureBox pic = sender as PictureBox;
            if (pic == null || selected.Contains(pic) || deck.Count == 0)
                return;

            string cardImg = deck[0];
            deck.RemoveAt(0);

            if (File.Exists(cardImg))
            {
                FlipCardWithAnimation(pic, cardImg);
                pic.Tag = cardImg;
                selected.Add(pic);

                if (tarotCardNames.TryGetValue(cardImg, out string cardName))
                {
                    cardToolTip.SetToolTip(pic, cardName);
                }
            }

            if (selected.Count == RequiredCardCount)
            {
                lblShow.Text = "抽牌完成";
                lblShow.BackColor = Color.Thistle;
                btnResetDraw.Enabled = true;
                btnViewResult.Enabled = true;

                var names = new List<string>();
                foreach (var p in selected)
                {
                    var path = p.Tag as string;
                    string name;
                    if (path != null && tarotCardNames.TryGetValue(path, out name))
                        names.Add(name);
                    else
                        names.Add("未知");
                }

                cardName1 = names.Count > 0 ? names[0] : "";
                cardName2 = names.Count > 1 ? names[1] : "";
                cardName3 = names.Count > 2 ? names[2] : "";

                foreach (var p in cardPics)
                {
                    p.Enabled = false;
                    if (!selected.Contains(p) && p.Image != null)
                        p.Image = ConvertToGrayscale(p.Image);
                }

                foreach (var p in selected)
                {
                    var path = p.Tag as string;
                    if (path != null && File.Exists(path))
                        p.Image = Image.FromFile(path);
                }
                // 是否塔羅
                if (isYesNoMode)
                {
                    YesNoselected = new List<PictureBox>(selected);
                    names = YesNoselected
                        .Select(pb => tarotCardNames[pb.Tag as string])
                        .ToList();
                    cardYesNoName1 = names.ElementAtOrDefault(0);
                    cardYesNoName2 = names.ElementAtOrDefault(1);
                    cardYesNoName3 = names.ElementAtOrDefault(2);
                }
            }
        }

        // 重新抽牌button
        private void btnResetDraw_Click(object sender, EventArgs e)
        {
            ResetDraw();
        }

        /* AI塔羅&是否塔羅結果 */
        private void tabPageResult_Enter(object sender, EventArgs e)
        {
            label48.BackColor = Color.LightCyan;
            label59.BackColor = Color.Transparent;
            checkBox2.BackColor = Color.Transparent;
        }
        private void tabPageYesNoResult_Enter(object sender, EventArgs e)
        {
            label55.BackColor = Color.AliceBlue;
            lblYesNoResult.BackColor = Color.AliceBlue;
            lblYesNoResult1.BackColor = Color.Transparent;
            lblYesNoResult2.BackColor = Color.Transparent;
            lblYesNoResult3.BackColor = Color.Transparent;
            label58.BackColor = Color.Transparent;
            checkBox1.BackColor = Color.Transparent;
        }
        
        private void btnViewResult_Click(object sender, EventArgs e)
        {
            string numText;
            if (isYesNoMode)
            {
                tabControlMain.SelectedTab = tabPageYesNoResult;

                picYesNoResult1.SizeMode = PictureBoxSizeMode.StretchImage;
                picYesNoResult2.SizeMode = PictureBoxSizeMode.StretchImage;
                picYesNoResult3.SizeMode = PictureBoxSizeMode.StretchImage;

                picYesNoResult1.Visible = drawCardCount == 3;
                picYesNoResult2.Visible = true;
                picYesNoResult3.Visible = drawCardCount == 3;
                void SetPic(PictureBox pic, int idx)
                {
                    var path = YesNoselected.ElementAtOrDefault(idx)?.Tag as string;
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        pic.Image = Image.FromFile(path);
                }
                if (drawCardCount == 1)
                    SetPic(picYesNoResult2, 0);
                else
                {
                    SetPic(picYesNoResult1, 0);
                    SetPic(picYesNoResult2, 1);
                    SetPic(picYesNoResult3, 2);
                }

                lblYesNoResult1.Visible = drawCardCount == 3;
                lblYesNoResult2.Visible = true;
                lblYesNoResult3.Visible = drawCardCount == 3;
                if (drawCardCount == 1)
                    lblYesNoResult2.Text = cardYesNoName1;
                else
                {
                    lblYesNoResult1.Text = cardYesNoName1;
                    lblYesNoResult2.Text = cardYesNoName2;
                    lblYesNoResult3.Text = cardYesNoName3;
                }

                // 準備 prompt
                numText = drawCardCount == 1 ? "單張" : "三張";
                string cardsTxt = drawCardCount == 1
                    ? $"「{cardYesNoName1}」"
                    : $"「{cardYesNoName1}」、「{cardYesNoName2}」、「{cardYesNoName3}」";
                string promptYesNo =
                    $"你是一位「療癒系塔羅占卜師，擅長用溫柔的語氣解釋牌義，會用比喻、鼓勵與祝福的方式說明」。\n" +
                    $"請使用「{numText}牌占卜」預測此問題「{userQuestion}」，" +
                    $"抽到的塔羅牌是{cardsTxt}，請根據此給我解讀。"+
                    $"第一行先溫柔的口吻只回答是或否或不確定其中一個詞";

                rtbYesNoAI.Text = "正在生成運勢解讀，請稍候...";
                var clientYesNo = new ChatClient(model: "gpt-4o", apiKey: apiKey);
                var resultYesNo = clientYesNo.CompleteChat(promptYesNo);      // ClientResult<ChatCompletion>
                var chatYesNo = resultYesNo.Value;                          // 取出 ChatCompletion

                
                rtbYesNoAI.Text = $"{chatYesNo.Content[0].Text}";
                return;
            }

            rtbAns.Text = "正在生成運勢解讀，請稍候...";
            tabControlMain.SelectedTab = tabPageResult;
            const string position = "正位";
            numText = RequiredCardCount == 1 ? "單張" : $"{RequiredCardCount}張";
            string prompt;
            if (RequiredCardCount == 1)
            {
                prompt =
                    $"你是一位「{style}」。\n" +
                    $"請使用「{numText}牌占卜」預測此問題「{userQuestion}」，" +
                    $"抽到的塔羅牌是「{cardName1}」{position}，請根據此牌給我解讀。";
            }
            else
            {
                prompt =
                    $"你是一位「{style}」。\n" +
                    $"請使用「{numText}牌占卜」預測此問題「{userQuestion}」，" +
                    $"依序抽到的塔羅牌是「{cardName1}」「{position}」、" +
                    $"「{cardName2}」「{position}」、" +
                    $"「{cardName3}」「{position}」，請根據這三張牌給我解讀。";
            }

            ChatClient client = new ChatClient(model: "gpt-4o", apiKey: apiKey);
            ChatCompletion completion = client.CompleteChat(prompt);

            
            rtbAns.Text = $"{completion.Content[0].Text}";

            textBox1.Visible = false;
            button1.Visible = false;
            label58.Visible = false;
        }
        public void SendCuteTarotEmail(string toEmail, string topic, string tarotMeaning)
        {
            string from = "liyingxiao992@gmail.com";
            string appPassword = "uqzt piqb balp gqpq";  // 建議用環境變數儲存
            string subject = "🔮 塔羅占卜結果 ✨";

            string htmlBody = $@"
            <html><body style='font-family:微軟正黑體'>
            <h2 style='color:#8A2BE2;'>🔮 占卜主題</h2>
            <p style='font-size:16px;'>{topic}</p>   
            <h2 style='color:#8A2BE2;'>🔮 占卜結果</h2>
            <p style='font-size:16px;'>{tarotMeaning}</p>
            <hr />
            <p style='color:#888;'>🌟 願宇宙給你溫柔的指引 ✨</p>
            <p>— 你的塔羅占卜小幫手 💜</p>
            </body></html>";

            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(from, "塔羅占卜機器人");
            mail.To.Add(toEmail);
            mail.Subject = subject;
            mail.Body = htmlBody;
            mail.IsBodyHtml = true;

            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            smtp.Credentials = new NetworkCredential(from, appPassword);
            smtp.EnableSsl = true;

            try
            {
                smtp.Send(mail);
                MessageBox.Show("💌 寄出成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("寄信失敗：" + ex.Message);
            }
        }
        //寄信--是否塔羅寄占卜結果
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = checkBox1.Checked;
            label58.Visible = isChecked;
            textBox1.Visible = isChecked;
            button1.Visible = isChecked;
        }
        private void button1_Click(object sender, EventArgs e)
        {

            string toEmail = textBox1.Text.Trim();
            string tarotResult = rtbYesNoAI.Text;

            if (string.IsNullOrEmpty(toEmail) || !toEmail.Contains("@"))
            {
                MessageBox.Show("請輸入正確的 Gmail！");
                return;
            }

            SendCuteTarotEmail(toEmail, userQuestion, tarotResult);  // 呼叫寄信方法
        }
        //寄信--AI塔羅寄占卜結果
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = checkBox2.Checked;
            label59.Visible = isChecked;
            textBox2.Visible = isChecked;
            button2.Visible = isChecked;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string toEmail = textBox2.Text.Trim();
            string tarotResult = rtbAns.Text;

            if (string.IsNullOrEmpty(toEmail) || !toEmail.Contains("@"))
            {
                MessageBox.Show("請輸入正確的 Gmail！");
                return;
            }

            SendCuteTarotEmail(toEmail, userQuestion, tarotResult);  // 呼叫寄信方法
        }
        // 洗牌
        private void Shuffle(List<string> list)
        {
            Random r = new Random();
            for (int i = list.Count - 1; i > 0; i--)
            {
                int j = r.Next(i + 1);
                (list[i], list[j]) = (list[j], list[i]);
            }
        }

        // 圖片轉灰階
        private Image ConvertToGrayscale(Image original)
        {
            if (original == null) return null;
            Bitmap grayBitmap = new Bitmap(original.Width, original.Height);
            using (Graphics g = Graphics.FromImage(grayBitmap))
            {
                ColorMatrix matrix = new ColorMatrix(new float[][]
                {
            new float[] { .3f, .3f, .3f, 0, 0 },
            new float[] { .59f, .59f, .59f, 0, 0 },
            new float[] { .11f, .11f, .11f, 0, 0 },
            new float[] { 0, 0, 0, 1, 0 },
            new float[] { 0, 0, 0, 0, 1 }
                });
                ImageAttributes attrs = new ImageAttributes();
                attrs.SetColorMatrix(matrix);
                g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height), 0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attrs);
            }
            return grayBitmap;
        }

        /* 塔羅運勢 */
        private string selectedFortuneType = "今日";
        private string selectedTimeRange = "";
        private int drawCardCountFortune = 3;
        private List<PictureBox> fortuneCardPics = new List<PictureBox>();
        private List<string> fortuneCardDeck = new List<string>();
        private List<PictureBox> fortuneSelected = new List<PictureBox>();
        private List<string> fortuneCardNames = new List<string>();
        private Dictionary<string, string> fortuneTarotNames = new Dictionary<string, string>
        {
            { "card_1.jpg", "愚人" },
            { "card_2.jpg", "魔術師" },
            { "card_3.jpg", "女祭司" },
            { "card_4.jpg", "皇后" },
            { "card_5.jpg", "皇帝" },
            { "card_6.jpg", "教宗" },
            { "card_7.jpg", "戀人" },
            { "card_8.jpg", "戰車" },
            { "card_9.jpg", "力量" },
            { "card_10.jpg", "隱士" },
            { "card_11.jpg", "命運之輪" },
            { "card_12.jpg", "正義" },
            { "card_13.jpg", "吊人" },
            { "card_14.jpg", "死神" },
            { "card_15.jpg", "節制" },
            { "card_16.jpg", "惡魔" },
            { "card_17.jpg", "塔" },
            { "card_18.jpg", "星星" },
            { "card_19.jpg", "月亮" },
            { "card_20.jpg", "太陽" },
            { "card_21.jpg", "審判" },
            { "card_22.jpg", "世界" }
        };

        private void tabPageTimeTarot_Enter(object sender, EventArgs e)
        {
            label15.BackColor = Color.LavenderBlush;
            groupBox1.BackColor = Color.LavenderBlush;
            groupBox2.BackColor = Color.LavenderBlush;
            label69.BackColor = Color.Transparent;
            label70.BackColor = Color.Transparent;
            label71.BackColor = Color.Transparent;
            label72.BackColor = Color.Transparent;
            label73.BackColor = Color.Transparent;
            rdbDailyFortune.BackColor = Color.Transparent;
            rdbMonthlyFortune.BackColor = Color.Transparent;
            rdbYearlyFortune.BackColor = Color.Transparent;
            rdbTarotA.BackColor = Color.Transparent;
            rdbTarotB.BackColor = Color.Transparent;
            rdbTarotC.BackColor = Color.Transparent;
            rdbTarotD.BackColor = Color.Transparent;
            rdbTarotE.BackColor = Color.Transparent;
            rdbTarotF.BackColor = Color.Transparent;

            // 預設選擇今日運勢
            rdbDailyFortune.Checked = true;
            selectedFortuneType = "今日";
            InitializeTarotReaderStyles();
        }

        private void InitializeTarotReaderStyles()
        {
            if (readerStyles.Count == 0)
            {
                readerStyles["reader1"] = "療癒系塔羅占卜師，擅長用溫柔的語氣解釋牌義，會用比喻、鼓勵與祝福的方式說明";
                readerStyles["reader2"] = "現實系塔羅占卜師，重視直觀與邏輯，解牌時會坦率地指出問題核心，提供具體可行的建議";
                readerStyles["reader3"] = "神秘系塔羅占卜師，注重牌與宇宙能量的連結，擅長從靈性視角詮釋命運的流動與內在課題";
                // 會員限定
                readerStyles["reader4"] = "心靈共鳴系塔羅占卜師，專注傾聽內心深層情感，透過塔羅釋放壓抑情緒，帶來情感療癒與自我成長";
                readerStyles["reader5"] = "未來趨勢系塔羅占卜師，結合占星能量與塔羅牌陣，剖析未來走向與關鍵時機，幫助掌握人生轉折與契機";
                readerStyles["reader6"] = "可愛治療系塔羅占卜師，擅長以柔和可愛的卡通符號與色彩，結合塔羅牌意象，輕鬆撫慰心靈";
            }

            rdbTarotA.Enabled = true;
            rdbTarotB.Enabled = true;
            rdbTarotC.Enabled = true;
            rdbTarotD.Enabled = true;
            rdbTarotE.Enabled = true;
            rdbTarotF.Enabled = true;

            rdbTarotA.Checked = true;
            selectedFortuneReaderStyle = readerStyles["reader1"];

            rdbTarotD.Enabled = isLoggedIn;
            rdbTarotE.Enabled = isLoggedIn;
            rdbTarotF.Enabled = isLoggedIn;

            // 非會員狀態下 radiobuttom 的顯示狀態
            if (!isLoggedIn)
            {
                rdbTarotD.ForeColor = Color.Gray;
                rdbTarotE.ForeColor = Color.Gray;
                rdbTarotF.ForeColor = Color.Gray;
            }
            else
            {
                // 會員狀態正常顯示
                rdbTarotD.ForeColor = SystemColors.ControlText;
                rdbTarotE.ForeColor = SystemColors.ControlText;
                rdbTarotF.ForeColor = SystemColors.ControlText;
                rdbTarotD.BackColor = SystemColors.Control;
                rdbTarotE.BackColor = SystemColors.Control;
                rdbTarotF.BackColor = SystemColors.Control;
            }
        }

        private string selectedFortuneReaderStyle = "";

        private void rdbTarotA_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbTarotA.Checked)
            {
                selectedFortuneReaderStyle = readerStyles["reader1"];
            }
        }

        private void rdbTarotB_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbTarotB.Checked)
            {
                selectedFortuneReaderStyle = readerStyles["reader2"];
            }
        }

        private void rdbTarotC_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbTarotC.Checked)
            {
                selectedFortuneReaderStyle = readerStyles["reader3"];
            }
        }

        private void rdbTarotD_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbTarotD.Checked)
            {
                selectedFortuneReaderStyle = readerStyles["reader4"];
            }
        }

        private void rdbTarotE_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbTarotE.Checked)
            {
                selectedFortuneReaderStyle = readerStyles["reader5"];
            }
        }

        private void rdbTarotF_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbTarotF.Checked)
            {
                selectedFortuneReaderStyle = readerStyles["reader6"];
            }
        }

        private void SetupFortuneReaderRadioButtons()
        {
            rdbTarotA.CheckedChanged += rdbTarotA_CheckedChanged;
            rdbTarotB.CheckedChanged += rdbTarotB_CheckedChanged;
            rdbTarotC.CheckedChanged += rdbTarotC_CheckedChanged;
            rdbTarotD.CheckedChanged += rdbTarotD_CheckedChanged;
            rdbTarotE.CheckedChanged += rdbTarotE_CheckedChanged;
            rdbTarotF.CheckedChanged += rdbTarotF_CheckedChanged;
        }

        private void rdbDailyFortune_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rdbDailyFortune.Checked)
            {
                selectedFortuneType = "今日";
            }
        }

        private void rdbMonthlyFortune_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rdbMonthlyFortune.Checked)
            {
                selectedFortuneType = "本月";
            }
        }

        private void rdbYearlyFortune_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rdbYearlyFortune.Checked)
            {
                selectedFortuneType = "年度";
            }
        }
        private void btnFortuneConfirm_Click_1(object sender, EventArgs e)
        {
            if (selectedFortuneType == "今日")
            {
                selectedTimeRange = "今日";
                drawCardCountFortune = 3; // 今日抽3
                tabControlMain.SelectedTab = tabPageFortuneDraw;
            }
            else if (selectedFortuneType == "本月")
            {
                // 本月運勢需要會員權限
                if (isLoggedIn)
                {
                    selectedTimeRange = "本月";
                    drawCardCountFortune = 4; // 本月抽4
                    tabControlMain.SelectedTab = tabPageFortuneDraw;
                }
                else
                {
                    ShowMembershipRequiredDialog(selectedFortuneType + "運勢");
                }
            }
            else if (selectedFortuneType == "年度")
            {
                // 年度運勢需要會員權限
                if (isLoggedIn)
                {
                    selectedTimeRange = "年度";
                    drawCardCountFortune = 6; // 年度抽6
                    tabControlMain.SelectedTab = tabPageFortuneDraw;
                }
                else
                {
                    ShowMembershipRequiredDialog(selectedFortuneType + "運勢");
                }
            }
        }

        // 跳出會員登入提示框
        private void ShowMembershipRequiredDialog(string featureName)
        {
            DialogResult result = MessageBox.Show(
                $"很抱歉，「{featureName}」功能僅限會員使用。\n登入會員即可享有更多功能！\n\n點擊「是」前往登入頁面，點擊「否」返回塔羅運勢",
                "會員專屬功能",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                // 選擇是，前往登入頁面
                tabControlMain.SelectedTab = tabPageLogin;
            }
            // 選擇否，則留在塔羅運勢頁面
        }

        /* 塔羅運勢抽牌 */
        private bool _fortuneFirstEnter = true;
        private void tabPageFortuneDraw_Enter_1(object sender, EventArgs e)
        {
            if (File.Exists("background.jpg"))
            {
                tabPageFortuneDraw.BackgroundImage = Image.FromFile("background.jpg");
                tabPageFortuneDraw.BackgroundImageLayout = ImageLayout.Stretch;
            }
            lblFortuneHint.BackColor = Color.Thistle;
            lblFortuneHint.Text = $"請為「{selectedTimeRange}運勢」選擇 {drawCardCountFortune} 張牌";
            if (_fortuneFirstEnter)
            {
                // 建立牌組
                fortuneCardPics = new List<PictureBox>
                {
                    picFortuneCard1, picFortuneCard2, picFortuneCard3, picFortuneCard4, picFortuneCard5, picFortuneCard6,
                    picFortuneCard7, picFortuneCard8, picFortuneCard9, picFortuneCard10, picFortuneCard11, picFortuneCard12,
                    picFortuneCard13, picFortuneCard14, picFortuneCard15, picFortuneCard16, picFortuneCard17, picFortuneCard18,
                    picFortuneCard19, picFortuneCard20, picFortuneCard21, picFortuneCard22
                };

                foreach (var pic in fortuneCardPics)
                {
                    pic.SizeMode = PictureBoxSizeMode.StretchImage;
                    pic.Click -= FortunePic_Click;
                    pic.Click += FortunePic_Click;
                    pic.Enabled = true;
                }
                if (originalPositions.Count == 0)
                {
                    foreach (var pic in fortuneCardPics)
                        originalPositions[pic] = pic.Location;
                }
                FortuneResetDraw();
                _fortuneFirstEnter = false;
            }
            lblFortuneHint.Text = $"請為「{selectedTimeRange}運勢」選擇 {drawCardCountFortune} 張牌";
        }
        private void FortuneResetDraw()
        {
            fortuneSelected.Clear();
            fortuneCardNames.Clear();
            fortuneCardDeck = Enumerable.Range(1, 22).Select(i => $"card_{i}.jpg").ToList();
            Shuffle(fortuneCardDeck);
            lblFortuneHint.BackColor = Color.Thistle;
            lblFortuneHint.Text = $"請為「{selectedTimeRange}運勢」選擇 {drawCardCountFortune} 張牌";

            btnFortuneResult.Enabled = false;
            btnFortuneRedraw.Enabled = false;

            foreach (var pic in fortuneCardPics)
            {
                if (File.Exists("card_cover.jpg"))
                {
                    pic.Image = Image.FromFile("card_cover.jpg");
                }
                pic.Enabled = true;
                pic.Tag = null;

                pic.MouseEnter += Pic_MouseEnter;
                pic.MouseLeave += Pic_MouseLeave;
            }
        }
        private void Pic_MouseEnter(object sender, EventArgs e)
        {
            PictureBox pic = sender as PictureBox;
            if (pic != null && pic.Enabled)
            {
                pic.BorderStyle = BorderStyle.Fixed3D;
                pic.BringToFront(); // 確保浮在上層

                if (originalPositions.TryGetValue(pic, out Point originalPos) &&
                    originalSizes.TryGetValue(pic, out Size originalSize))
                {
                    int newWidth = (int)(originalSize.Width * scaleFactor);
                    int newHeight = (int)(originalSize.Height * scaleFactor);

                    // 計算新的位置以保持中心對齊
                    int offsetX = (newWidth - originalSize.Width) / 2;
                    int offsetY = (newHeight - originalSize.Height) / 2;

                    pic.Size = new Size(newWidth, newHeight);
                    pic.Location = new Point(originalPos.X - offsetX, originalPos.Y - offsetY);
                }
            }
        }

        private void Pic_MouseLeave(object sender, EventArgs e)
        {
            PictureBox pic = sender as PictureBox;
            if (pic != null)
            {
                pic.BorderStyle = BorderStyle.None;

                if (originalPositions.TryGetValue(pic, out Point originalPos) &&
                    originalSizes.TryGetValue(pic, out Size originalSize))
                {
                    pic.Size = originalSize;
                    pic.Location = originalPos;
                }
            }
        }
        private void FortunePic_Click(object sender, EventArgs e)
        {
            if (fortuneSelected.Count >= drawCardCountFortune) return;
            PictureBox pic = sender as PictureBox;
            if (pic == null || fortuneSelected.Contains(pic)) return;
            if (fortuneCardDeck.Count == 0) return;

            string cardImg = fortuneCardDeck[0];
            fortuneCardDeck.RemoveAt(0);

            if (File.Exists(cardImg))
            {
                FlipCardWithAnimation(pic, cardImg);
                pic.Tag = cardImg;
                fortuneSelected.Add(pic);

                if (tarotCardNames.TryGetValue(cardImg, out var cardName))
                {
                    cardToolTip.SetToolTip(pic, cardName);
                }

                if (fortuneTarotNames.ContainsKey(cardImg))
                {
                    fortuneCardNames.Add(fortuneTarotNames[cardImg]);
                }
                else
                {
                    fortuneCardNames.Add("未知");
                }
            }

            if (fortuneSelected.Count == drawCardCountFortune)
            {
                lblFortuneHint.BackColor = Color.Thistle;
                lblFortuneHint.Text = "抽牌完成";
                btnFortuneRedraw.Enabled = true;
                btnFortuneResult.Enabled = true;

                foreach (var p in fortuneCardPics)
                {
                    p.Enabled = false;
                    if (!fortuneSelected.Contains(p) && p.Image != null)
                        p.Image = ConvertToGrayscale(p.Image);
                }
            }
        }
        private void btnFortuneRedraw_Click_1(object sender, EventArgs e)
        {
            FortuneResetDraw();
        }

        private void btnFortuneResult_Click_1(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageFortuneResult;
            DisplayFortuneCards();
            GenerateFortuneResult();
        }

        /* 塔羅運勢結果 */
        private void tabPageFortuneResult_Enter(object sender, EventArgs e)
        {
            label53.BackColor = Color.LavenderBlush;
            lblFortuneResult1.BackColor = Color.Transparent;
            lblFortuneResult2.BackColor = Color.Transparent;
            lblFortuneResult3.BackColor = Color.Transparent;
            lblFortuneResult4.BackColor = Color.Transparent;
            lblFortuneResult5.BackColor = Color.Transparent;
            lblFortuneResult6.BackColor = Color.Transparent;
            label60.BackColor = Color.Transparent;
            checkBox3.BackColor = Color.Transparent;
        }

        // 展示抽到的牌
        private void DisplayFortuneCards()
        {
            for (int i = 1; i <= 6; i++)
            {
                Control lblControl = this.Controls.Find($"lblFortuneResult{i}", true).FirstOrDefault();
                Control picControl = this.Controls.Find($"picFortuneResult{i}", true).FirstOrDefault();

                if (lblControl != null)
                {
                    lblControl.Visible = false;
                }
                if (picControl != null)
                {
                    picControl.Visible = false;
                }
            }

            // 根據抽到的牌數顯示對應數量的 label 和 PictureBox
            for (int i = 0; i < fortuneSelected.Count; i++)
            {
                int controlIndex = i + 1;

                Label lblResult = this.Controls.Find($"lblFortuneResult{controlIndex}", true).FirstOrDefault() as Label;
                PictureBox picResult = this.Controls.Find($"picFortuneResult{controlIndex}", true).FirstOrDefault() as PictureBox;

                if (lblResult != null && picResult != null)
                {
                    lblResult.Visible = true;
                    picResult.Visible = true;

                    if (i < fortuneCardNames.Count)
                    {
                        lblResult.Text = fortuneCardNames[i];
                    }
                    else
                    {
                        lblResult.Text = "未知";
                    }

                    string cardImg = fortuneSelected[i].Tag?.ToString();
                    if (!string.IsNullOrEmpty(cardImg) && File.Exists(cardImg))
                    {
                        picResult.SizeMode = PictureBoxSizeMode.StretchImage;
                        picResult.Image = Image.FromFile(cardImg);
                    }
                }
            }
        }

        // 塔羅運勢模型生成結果
        private void GenerateFortuneResult()
        {
            string style = selectedFortuneReaderStyle;
            string cardNamesString = string.Join("」、「", fortuneCardNames);
            string prompt = $"你是一位{style}。\n請根據所選擇的{drawCardCountFortune}張牌，預測「{selectedTimeRange}」運勢。\n" +
                           $"抽到的塔羅牌是「{cardNamesString}」。\n" +
                           $"請詳細解讀這些牌面對於{selectedTimeRange}運勢的含義，可以包括課業、感情或財運，和整體運勢走向。";

            rtbFortuneAI.ReadOnly = true;
            rtbFortuneAI.Text = "正在生成運勢解讀，請稍候...";
            Application.DoEvents();

            ChatClient client = new ChatClient(model: "gpt-4o", apiKey: apiKey);
            ChatCompletion completion = client.CompleteChat(prompt);

            rtbFortuneAI.Text = $"解析：{completion.Content[0].Text}";
        }
        //寄信--塔羅運勢
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = checkBox3.Checked;
            label60.Visible = isChecked;
            textBox3.Visible = isChecked;
            button3.Visible = isChecked;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string toEmail = textBox3.Text.Trim();
            string tarotResult = rtbFortuneAI.Text;
            string userQuestion = $"預測{ selectedTimeRange}運勢";
            if (string.IsNullOrEmpty(toEmail) || !toEmail.Contains("@"))
            {
                MessageBox.Show("請輸入正確的 Gmail！");
                return;
            }

            SendCuteTarotEmail(toEmail, userQuestion, tarotResult);  // 呼叫寄信方法
        }

        /* 會員中心 */
        private bool isEditingMember = false;
        private string currentUserAccount = "";
        private Dictionary<string, string> originalMemberData = new Dictionary<string, string>();

        private void tabPageMemberCenter_Enter(object sender, EventArgs e)
        {
            label32.BackColor = Color.Transparent;
            label35.BackColor = Color.Transparent;
            label36.BackColor = Color.Transparent;
            label37.BackColor = Color.Transparent;
            label38.BackColor = Color.Transparent;
            label39.BackColor = Color.Transparent;
            label40.BackColor = Color.Transparent;
            label43.BackColor = Color.Transparent;
            rdoMemberMale.BackColor = Color.Transparent;
            rdoMemberFemale.BackColor = Color.Transparent;

            if (!File.Exists("users.csv") || string.IsNullOrEmpty(currentUserAccount)) return;

            var line = File.ReadAllLines("users.csv", Encoding.UTF8)
                           .Skip(1)
                           .FirstOrDefault(l => l.Split(',')[1] == currentUserAccount);
            if (line == null) return;

            var parts = line.Split(',');
            txtMemberNickname.Text = parts[0];
            txtMemberAccount.Text = parts[1];
            txtMemberPassword.Text = parts[2];
            rdoMemberMale.Checked = parts[3] == "生理男";
            rdoMemberFemale.Checked = parts[3] == "生理女";
            if (DateTime.TryParse(parts[4], out DateTime bd)) dtpMemberBirthday.Value = bd;
            txtMemberEmail.Text = parts[5];

            // 備份
            originalMemberData["暱稱"] = parts[0];
            originalMemberData["密碼"] = parts[2];
            originalMemberData["性別"] = parts[3];
            originalMemberData["生日"] = parts[4];
            originalMemberData["Email"] = parts[5];

            SetMemberEditMode(false);
            btnMemberEditSave.Text = "修改資料";
            isEditingMember = false;
        }

        private void tabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControlMain.SelectedTab == tabPageHome)
            {
                swimming.Play();
            }

            if (tabControlMain.SelectedTab == tabPageAITarot)
            {
                killswitch.Play();
            }

            if (tabControlMain.SelectedTab == tabPageYesNoTarot)
            {
                Beginning.Play();
                btnThreeTarot.Enabled = isLoggedIn;
            }

            if (tabControlMain.SelectedTab == tabPageTimeTarot)
            {
                Immaterial.Play();
            }

            if (tabControlMain.SelectedTab == tabPageAppointment)
            {
                dream.Play();
            }
        }

        // 控制欄位是否可編輯
        private void SetMemberEditMode(bool editable)
        {
            txtMemberAccount.ReadOnly = true; // 始終唯讀
            txtMemberNickname.ReadOnly = !editable;
            txtMemberPassword.ReadOnly = !editable;
            txtMemberEmail.ReadOnly = !editable;
            dtpMemberBirthday.Enabled = editable;
            rdoMemberMale.Enabled = editable;
            rdoMemberFemale.Enabled = editable;
            isEditingMember = editable;
        }

        // 修改資料／儲存資料按鈕
        private void btnMemberEditSave_Click(object sender, EventArgs e)
        {
            if (!isEditingMember)
            {
                // 切換到編輯模式
                SetMemberEditMode(true);
                btnMemberEditSave.Text = "儲存資料";
                isEditingMember = true;
                return;
            }

            // 儲存流程
            string newNick = txtMemberNickname.Text.Trim();
            string newPwd = txtMemberPassword.Text.Trim();
            string newGen = rdoMemberMale.Checked ? "生理男" : "生理女";
            string newBday = dtpMemberBirthday.Value.ToString("yyyy/MM/dd");
            string newEmail = txtMemberEmail.Text.Trim();

            // 驗證
            if (string.IsNullOrEmpty(newPwd) ||
                !Regex.IsMatch(newPwd, @"^[\p{L}\p{N}\p{P}]{1,20}$"))
            {
                MessageBox.Show("密碼必填，且僅限 1~20 字元，英數字與符號組合");
                return;
            }
            if (string.IsNullOrEmpty(newNick))
            {
                MessageBox.Show("暱稱必填，且長度不得超過10字元");
                return;
            }
            var allNicks = File.ReadAllLines("users.csv", Encoding.UTF8)
                               .Skip(1)
                               .Select(l => l.Split(',')[0]);
            if (newNick != originalMemberData["暱稱"] && allNicks.Contains(newNick))
            {
                MessageBox.Show($"暱稱「{newNick}」已被使用，請更換其他暱稱！");
                return;
            }
            if (string.IsNullOrEmpty(newEmail) ||
                !Regex.IsMatch(newEmail, @"^[\w\.\+-]+@[\w-]+\.[\w\.-]+$"))
            {
                MessageBox.Show("Email 格式不正確，請輸入正確的 Email 格式");
                return;
            }

            // 比對哪些欄位改變
            var modified = new List<string>();
            if (originalMemberData["暱稱"] != newNick) modified.Add("暱稱");
            if (originalMemberData["密碼"] != newPwd) modified.Add("密碼");
            if (originalMemberData["性別"] != newGen) modified.Add("性別");
            if (originalMemberData["生日"] != newBday) modified.Add("生日");
            if (originalMemberData["Email"] != newEmail) modified.Add("Email");

            if (modified.Count == 0)
            {
                MessageBox.Show("資料無異動");
            }
            else
            {
                // 更新導覽列歡迎文字
                loggedInUser = newNick;
                navBar.SetLoginState(true, newNick);

                // 寫回 CSV
                var lines = File.ReadAllLines("users.csv", Encoding.UTF8).ToList();
                for (int i = 1; i < lines.Count; i++)
                {
                    var p = lines[i].Split(',');
                    if (p[1] == currentUserAccount)
                    {
                        p[0] = newNick; p[2] = newPwd; p[3] = newGen; p[4] = newBday; p[5] = newEmail;
                        lines[i] = string.Join(",", p);
                        break;
                    }
                }
                File.WriteAllLines("users.csv", lines, new UTF8Encoding(true));
                MessageBox.Show($"以下資料已修改：{string.Join("、", modified)}");
            }

            // 回唯讀
            SetMemberEditMode(false);
            btnMemberEditSave.Text = "修改資料";
            isEditingMember = false;

            // 更新備份
            originalMemberData["暱稱"] = newNick;
            originalMemberData["密碼"] = newPwd;
            originalMemberData["性別"] = newGen;
            originalMemberData["生日"] = newBday;
            originalMemberData["Email"] = newEmail;
        }

        private bool ConfirmLeaveMemberCenter()
        {
            if (tabControlMain.SelectedTab == tabPageMemberCenter && isEditingMember)
            {
                var r = MessageBox.Show(
                  "尚未儲存的修改會遺失，確定要離開嗎？",
                  "尚未儲存提醒",
                  MessageBoxButtons.YesNo,
                  MessageBoxIcon.Warning);
                if (r == DialogResult.No) return false;

                // 使用者確定放棄編輯 → 回復唯讀
                SetMemberEditMode(false);
                btnMemberEditSave.Text = "修改資料";
                isEditingMember = false;
            }
            return true;
        }

        private void TabControlMain_Selecting(object sender, TabControlCancelEventArgs e)
        {
            // 如果當前分頁是會員中心，而且正在編輯
            if (tabControlMain.SelectedTab == tabPageMemberCenter && isEditingMember)
            {
                var result = MessageBox.Show(
                    "尚未儲存的修改會遺失，確定要離開嗎？",
                    "尚未儲存提醒",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (result == DialogResult.No)
                {
                    // 取消本次切換
                    e.Cancel = true;
                }
                else
                {
                    // 使用者確認放棄編輯 → 回復唯讀
                    SetMemberEditMode(false);
                    btnMemberEditSave.Text = "修改資料";
                    isEditingMember = false;
                }
            }
        }

        /* 預約占卜 */
        private readonly string[] _readerCodes = { "A", "B", "C", "D" };
        private readonly string[] _readerNames = { "子嫣", "秉廉", "禮英", "煜瑋" };
        private DataTable scheduleRaw;
        private void tabPageAppointment_Enter(object sender, EventArgs e)
        {
            label62.BackColor = Color.Transparent;
            label74.BackColor = Color.Transparent;
            label63.BackColor = Color.Transparent;
            label64.BackColor = Color.Transparent;
            label67.BackColor = Color.Transparent;
            label33.BackColor = Color.Transparent;


            cmbAppointmentReader.Items.Clear();
            cmbAppointmentReader.Items.AddRange(_readerNames);
            cmbAppointmentReader.SelectedIndexChanged -= cmbAppointmentReader_SelectedIndexChanged;
            cmbAppointmentReader.SelectedIndexChanged += cmbAppointmentReader_SelectedIndexChanged;
            cmbAppointmentReader.SelectedIndex = 0;

            LoadAppointmentSchedule();
        }

        private void cmbAppointmentReader_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadAppointmentSchedule();
        }

        private void LoadAppointmentSchedule()
        {
            string path = Path.Combine(Application.StartupPath, "schedules.csv");
            if (!File.Exists(path))
            {
                dgvAppointmentSchedule.DataSource = null;
                scheduleRaw = null;
                return;
            }

            // 1. 先把原始表全讀進 dtRaw
            var dtRaw = new DataTable();
            using (var sr = new StreamReader(path, Encoding.UTF8))
            {
                // header
                var headers = SplitCsvLine(sr.ReadLine());
                foreach (var h in headers)
                    dtRaw.Columns.Add(h);

                // 每一行
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    var fields = SplitCsvLine(line);
                    if (fields.Length < dtRaw.Columns.Count)
                    {
                        // 不夠就補空串
                        fields = fields
                            .Concat(Enumerable.Repeat(string.Empty, dtRaw.Columns.Count - fields.Length))
                            .ToArray();
                    }
                    else if (fields.Length > dtRaw.Columns.Count)
                    {
                        // 多了就截斷
                        fields = fields.Take(dtRaw.Columns.Count).ToArray();
                    }
                    dtRaw.Rows.Add(fields);
                }
            }
            // 存起來以便後來更新寫回
            scheduleRaw = dtRaw;

            // 2. 用一份 Clone() 來顯示
            var dt = dtRaw.Copy();

            // 3. 根據選取的塔羅師 code 映射「可/不可預約」
            string code = _readerCodes[cmbAppointmentReader.SelectedIndex];
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 1; c < dt.Columns.Count; c++)
                {
                    var cell = Convert.ToString(dt.Rows[r][c]);
                    var slots = cell.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(x => x.Trim());
                    bool available = slots.Any(x =>
                        x.Equals(code, StringComparison.OrdinalIgnoreCase));
                    dt.Rows[r][c] = available ? "可預約" : "不可預約";
                }
            }

            // 4. 綁定到 DataGridView
            dgvAppointmentSchedule.DataSource = dt;
            dgvAppointmentSchedule.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.AllCells;
            dgvAppointmentSchedule.Columns[0].HeaderText = "";
            dgvAppointmentSchedule.ReadOnly = true;
            dgvAppointmentSchedule.AllowUserToAddRows = false;
            dgvAppointmentSchedule.AllowUserToDeleteRows = false;
            dgvAppointmentSchedule.AllowUserToResizeColumns = false;
            dgvAppointmentSchedule.AllowUserToResizeRows = false;
            dgvAppointmentSchedule.ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dgvAppointmentSchedule.RowHeadersWidthSizeMode =
                DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dgvAppointmentSchedule.RowHeadersVisible = false;
        }

        private string[] SplitCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var sb = new StringBuilder();
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"'); i++;
                    }
                    else inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result.ToArray();
        }


        private void dgvAppointmentSchedule_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 1) return;
            var date = dgvAppointmentSchedule.Rows[e.RowIndex].Cells[0].Value.ToString();
            var time = dgvAppointmentSchedule.Columns[e.ColumnIndex].HeaderText;
            var readerName = cmbAppointmentReader.SelectedItem.ToString();
            var code = _readerCodes[cmbAppointmentReader.SelectedIndex];
            // 未登入
            if (!isLoggedIn)
            {
                var dr = MessageBox.Show("只有會員才能預約，是否登入？", "登入提醒",
                                         MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr == DialogResult.Yes) tabControlMain.SelectedTab = tabPageLogin;
                return;
            }
            // 只能預約明天以後
            if (DateTime.TryParse(date, out var d0) && d0 < DateTime.Today.AddDays(1))
            {
                MessageBox.Show("只能預約明天以後的時段！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // 不可預約
            if (dgvAppointmentSchedule.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "不可預約")
            {
                MessageBox.Show("此時段無法預約", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 確認
            if (MessageBox.Show($"確定要預約 {readerName}，{date}，{time}？", "確認",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                != DialogResult.Yes) return;

            {
                var rawCell = Convert.ToString(scheduleRaw.Rows[e.RowIndex][e.ColumnIndex]);
                var list = rawCell.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                  .Select(s => s.Trim())
                                  .Where(s => !s.Equals(code, StringComparison.OrdinalIgnoreCase))
                                  .ToArray();
                scheduleRaw.Rows[e.RowIndex][e.ColumnIndex] = string.Join(",", list);
                SaveSchedulesCsv();
            }

            AppendAppointmentInfo(date, time, readerName);
            LoadAppointmentSchedule();
            MessageBox.Show("預約成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SaveSchedulesCsv()
        {
            if (scheduleRaw == null) return;
            string path = Path.Combine(Application.StartupPath, "schedules.csv");

            string Quote(string field)
            {
                if (field.Contains("\""))
                    field = field.Replace("\"", "\"\"");
                if (field.Contains(",") || field.Contains("\"") || field.Contains("\r") || field.Contains("\n"))
                    return $"\"{field}\"";
                return field;
            }

            using (var sw = new StreamWriter(path, false, Encoding.UTF8))
            {
                var headers = scheduleRaw.Columns.Cast<DataColumn>()
                                     .Select(c => c.ColumnName)
                                     .ToList();
                var timeSlots = headers.Skip(1).Select(Quote);
                sw.WriteLine("," + string.Join(",", timeSlots));

                foreach (DataRow row in scheduleRaw.Rows)
                {
                    var cells = new List<string>
            {
                row[0].ToString()  // 日期
            };
                    for (int c = 1; c < scheduleRaw.Columns.Count; c++)
                    {
                        cells.Add(Quote(row[c].ToString()));
                    }
                    sw.WriteLine(string.Join(",", cells));
                }
            }
        }

        private void AppendAppointmentInfo(string date, string time, string readerName)
        {
            string file = Path.Combine(Application.StartupPath, "appointments.csv");
            var lines = File.Exists(file)
                ? File.ReadAllLines(file, Encoding.UTF8).ToList()
                : new List<string> { "日期,時間,塔羅師,暱稱,帳號,性別,生日,Email" };

            // 抓會員資料
            var uline = File.ReadAllLines(Path.Combine(Application.StartupPath, "users.csv"), Encoding.UTF8)
                            .Skip(1)
                            .FirstOrDefault(l => l.Split(',')[1] == loggedInAccount);
            if (uline == null) return;
            var u = uline.Split(',');
            lines.Add($"{date},{time},{readerName},{u[0]},{u[1]},{u[3]},{u[4]},{u[5]}");

            // 排序
            var header = lines[0];
            var data = lines.Skip(1)
                .Select(l =>
                {
                    var f = l.Split(',');
                    return new
                    {
                        dt = DateTime.Parse(f[0]),
                        ts = TimeSpan.Parse(f[1].Split('-')[0]),
                        rd = f[2],
                        line = l
                    };
                })
                .OrderBy(x => x.dt)
                .ThenBy(x => x.ts)
                .ThenBy(x => x.rd)
                .Select(x => x.line)
                .ToList();

            using (var sw = new StreamWriter(file, false, Encoding.UTF8))
            {
                sw.WriteLine(header);
                data.ForEach(l => sw.WriteLine(l));
            }
        }

        /* 預約後台 */
        private DataTable dtAppointments;
        private void FormatGrid(DataGridView dgv)
        {
            dgv.ReadOnly = true;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowUserToResizeColumns = false;
            dgv.AllowUserToResizeRows = false;
            dgv.RowHeadersVisible = false;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }
        private void LoadAllAppointments()
        {
            string path = Path.Combine(Application.StartupPath, "appointments.csv");
            dtAppointments = new DataTable();
            if (!File.Exists(path)) return;

            using (var sr = new StreamReader(path, Encoding.UTF8))
            {
                // header
                var cols = sr.ReadLine().Split(',');
                foreach (var c in cols) dtAppointments.Columns.Add(c);

                // rows
                string line;
                while ((line = sr.ReadLine()) != null)
                    dtAppointments.Rows.Add(line.Split(','));
            }
        }

        private void tabPageAppointmentAdmin_Enter(object sender, EventArgs e)
        {
            label65.BackColor = Color.Transparent;
            label66.BackColor = Color.Transparent;
            rbByDate.BackColor = Color.Transparent;
            rbByReader.BackColor = Color.Transparent;

            tabPageAppointmentAdmin.Enter += tabPageAppointmentAdmin_Enter;
            rbByDate.CheckedChanged += ModeChanged;
            rbByReader.CheckedChanged += ModeChanged;
            cmbFilterReader.SelectedIndexChanged += cmbFilterReader_SelectedIndexChanged;
            dgvSummary.SelectionChanged += dgvSummary_SelectionChanged;
            LoadAllAppointments();

            // 初始化下拉：在「依塔羅師檢視」才用到
            cmbFilterReader.Items.Clear();
            cmbFilterReader.Items.AddRange(_readerNames);

            // 預設模式
            rbByDate.Checked = true;
        }
        private void ModeChanged(object sender, EventArgs e)
        {
            // 依塔羅師檢視時才啟用下拉
            cmbFilterReader.Enabled = rbByReader.Checked;

            if (rbByDate.Checked)
                ShowSummaryByDate();
            else
                ShowSummaryByReader();
        }
        private void ShowSummaryByDate()
        {
            // 彙總：每個日期的筆數
            var summary = dtAppointments.AsEnumerable()
                .GroupBy(r => r.Field<string>("日期"))
                .Select(g => new {
                    日期 = g.Key,
                    筆數 = g.Count()
                })
                .OrderBy(x => DateTime.Parse(x.日期))
                .ToList();

            // 建 DataTable
            var dt = new DataTable();
            dt.Columns.Add("日期");
            dt.Columns.Add("預約筆數", typeof(int));
            summary.ForEach(x => dt.Rows.Add(x.日期, x.筆數));

            // 載入 Grid
            dgvSummary.DataSource = dt;
            FormatGrid(dgvSummary);

            // 選第一列以觸發 SelectionChanged
            if (dgvSummary.Rows.Count > 0)
                dgvSummary.Rows[0].Selected = true;
        }

        private void dgvSummary_SelectionChanged(object sender, EventArgs e)
        {
            // 先清空 Details
            dgvDetails.DataSource = null;

            // 依「日期」檢視
            if (rbByDate.Checked)
            {
                if (dgvSummary.SelectedRows.Count == 0) return;
                string selDate = dgvSummary.SelectedRows[0].Cells["日期"].Value.ToString();
                var rows = dtAppointments.AsEnumerable()
                    .Where(r => r.Field<string>("日期") == selDate)
                    .CopyToDataTableOrEmpty();
                dgvDetails.DataSource = rows;
                FormatGrid(dgvDetails);
            }
            // 依「塔羅師」檢視：點 Summary 直接切換
            else if (rbByReader.Checked)
            {
                if (dgvSummary.SelectedRows.Count == 0) return;
                string selReader = dgvSummary.SelectedRows[0].Cells["塔羅師"].Value.ToString();
                // 同樣更新下拉的選項（保持同步）
                cmbFilterReader.SelectedIndexChanged -= cmbFilterReader_SelectedIndexChanged;
                cmbFilterReader.SelectedItem = selReader;
                cmbFilterReader.SelectedIndexChanged += cmbFilterReader_SelectedIndexChanged;

                // 篩出該師的所有預約
                var rows = dtAppointments.AsEnumerable()
                    .Where(r => r.Field<string>("塔羅師") == selReader)
                    .CopyToDataTableOrEmpty();
                dgvDetails.DataSource = rows;
                FormatGrid(dgvDetails);
            }
        }

        private void ShowSummaryByReader()
        {
            // 彙總：每位師的預約筆數
            var summary = dtAppointments.AsEnumerable()
                .GroupBy(r => r.Field<string>("塔羅師"))
                .Select(g => new {
                    塔羅師 = g.Key,
                    筆數 = g.Count()
                })
                .OrderBy(x => x.塔羅師)
                .ToList();

            // 建 DataTable
            var dt = new DataTable();
            dt.Columns.Add("塔羅師");
            dt.Columns.Add("預約筆數", typeof(int));
            summary.ForEach(x => dt.Rows.Add(x.塔羅師, x.筆數));

            // 載入 Summary 並格式化
            dgvSummary.DataSource = dt;
            FormatGrid(dgvSummary);

            // 讓欄寬填滿整個表格
            dgvSummary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // 如果有列，選第一筆，以便一進來就有明細
            if (dgvSummary.Rows.Count > 0)
                dgvSummary.Rows[0].Selected = true;
            else
                dgvDetails.DataSource = null;
        }

        private void cmbFilterReader_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!rbByReader.Checked) return;

            string selReader = cmbFilterReader.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selReader))
            {
                dgvDetails.DataSource = null;
            }
            else
            {
                // 把 summary 中對應列選起來
                for (int i = 0; i < dgvSummary.Rows.Count; i++)
                {
                    if ((string)dgvSummary.Rows[i].Cells["塔羅師"].Value == selReader)
                    {
                        dgvSummary.Rows[i].Selected = true;
                        dgvSummary.FirstDisplayedScrollingRowIndex = i;
                        break;
                    }
                }

                // 篩 Details
                var rows = dtAppointments.AsEnumerable()
                    .Where(r => r.Field<string>("塔羅師") == selReader)
                    .CopyToDataTableOrEmpty();
                dgvDetails.DataSource = rows;
                FormatGrid(dgvDetails);
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            var dlg = new SaveFileDialog();
            dlg.Filter = "CSV 檔案 (*.csv)|*.csv";
            dlg.FileName = "appointments_export.csv";
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            var dt = dtAppointments;
            var sw = new StreamWriter(dlg.FileName, false, Encoding.UTF8);

            var headers = dt.Columns.Cast<DataColumn>()
                            .Select(c => c.ColumnName);
            sw.WriteLine(string.Join(",", headers));

            foreach (DataRow row in dt.Rows)
            {
                var fields = row.ItemArray
                                .Select(field => field?.ToString().Replace(",", "，"));
                sw.WriteLine(string.Join(",", fields));
            }

            sw.Close();
            MessageBox.Show("資料已匯出 ", "匯出完成",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnShowChart_Click(object sender, EventArgs e)
        {
            var dt = dtAppointments;

            chartAnalysis.Series.Clear();
            chartAnalysis.ChartAreas.Clear();
            chartAnalysis.ChartAreas.Add("Default");

            var cntByTime = dt.AsEnumerable()
                .GroupBy(r => r.Field<string>("時間"))
                .Select(g => new {
                    TimeSlot = g.Key,
                    Count = g.Count()
                })
                .OrderBy(x => x.TimeSlot)
                .ToList();

            var ser = new System.Windows.Forms.DataVisualization.Charting.Series("預約數");
            ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            chartAnalysis.Series.Add(ser);

            foreach (var item in cntByTime)
            {
                ser.Points.AddXY(item.TimeSlot, item.Count);
            }

            chartAnalysis.Visible = true;
        }
    }
    public static class LinqExtensions
    {
        public static DataTable CopyToDataTableOrEmpty(this IEnumerable<DataRow> source)
        {
            var list = source.ToList();
            if (list.Count == 0)
                return new DataTable();  // 空的、沒欄位
            else
                return list.CopyToDataTable();
        }
    }
}
