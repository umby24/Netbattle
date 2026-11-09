using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Network;
using Netbattle.Network;
using System.Runtime.InteropServices;
using Netbattle.Common;

namespace Netbattle.Forms {
    // -- Port of Player.frm.
    public partial class ChallengeWindow : Form
    {
        /// <summary>
        /// The open challenge window, if any. VB only ever had one.
        /// </summary>
        public static ChallengeWindow? Current { get; private set; }

        // -- State Flags.
        public bool ICalled;
        public bool PressedOk;
        public int GameMode;
        private readonly NbClient _client;
        private readonly Player _challenger;
        private readonly ChallengeSettings _settings; // -- What we were challenged with, when !ICalled.

        private readonly bool[] _canBattle = new bool[5]; // -- Indexed by mode, 1 to 4.
        private readonly int[] _randPoke = new int[7];    // -- 1 to 6, the Challenge Cup preview.
        private int _randStep;
        private readonly HashSet<int> _ghosted = new HashSet<int>(); // -- Rules that can't be changed.
        private readonly Random _random = new Random();
        private bool _replied; // -- We've answered, or the server ended it, so closing mustn't refuse/cancel again.

        private static readonly Color GhostColor = Color.FromArgb(161, 161, 161);

        /// <summary>
        /// Used when viewing / challenging someone else.
        /// </summary>
        public ChallengeWindow(NbClient nbClient, Player challenger, Form parent)
        {
            this.MdiParent = parent.MdiParent;
            InitializeComponent();
            _client = nbClient;
            _challenger = challenger;
            _settings = new ChallengeSettings { Rules = new bool[ChallengeRules.RuleCount + 1] };
            ICalled = true;
            Init();
        }

        /// <summary>
        /// Used when someone actually challenges us
        /// </summary>
        public ChallengeWindow(NbClient nbClient, Player challenger, Form parent, ChallengeSettings settings)
        {
            this.MdiParent = parent.MdiParent;
            InitializeComponent();
            _client = nbClient;
            _challenger = challenger;
            _settings = settings;
            GameMode = settings.Mode;
            Init();
        }

        private void Init()
        {
            this.Top = (Screen.PrimaryScreen.Bounds.Height - this.Height) / 2;
            Current = this;

            _client.ChallengeNumber = _challenger.Id;
            _client.ChallengeIsOutgoing = ICalled;
            _client.ChallengeWindowOpen = true;
            _client.ChallengeEnded += ClientOnChallengeEnded;
        }

        /// <summary>
        /// Closes the window the way the Cancel / Refuse button does, telling the other player.
        /// </summary>
        public void CloseAndCancel()
        {
            btnCancel_Click(this, EventArgs.Empty);
        }

        private void ChallengeWindow_Load(object sender, EventArgs e)
        {
            PopulateForm();
        }

        private void PopulateForm()
        {
            Player me = _client.Me;

            // -- Figure out which modes are supported & fill in the list
            int listType = 3;
            BattleModes g1 = ChallengeRules.CompatVersion(me.GameVersion);
            BattleModes g2 = ChallengeRules.CompatVersion(_challenger.GameVersion);
            if (g1 == BattleModes.nbRBYBattle && g2 == BattleModes.nbRBYBattle)
            {
                _canBattle[1] = true;
                _canBattle[2] = true;
                listType = 0;
            }
            else if (g1 != BattleModes.nbAdvBattle && g2 != BattleModes.nbAdvBattle)
            {
                _canBattle[2] = true;
                listType = 1;
                if (RbyCompatible(me) && RbyCompatible(_challenger))
                {
                    _canBattle[1] = true;
                    listType = 0;
                }
            }
            else if (g1 == BattleModes.nbAdvBattle && g2 == BattleModes.nbAdvBattle)
            {
                _canBattle[3] = true;
                _canBattle[4] = true;
                listType = 2;
            }

            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(ChallengeRules.ModeText);
            comboBox1.SelectedIndex = listType;

            // -- Fill in terrain list. Its index is the terrain number sent: 0 is random.
            comboBox2.Items.Clear();
            comboBox2.Items.Add("Random Terrain");
            foreach (TerrainTypes terrain in Enum.GetValues<TerrainTypes>())
            {
                string text = ChallengeRules.TerrainText(terrain);
                if (text != "")
                    comboBox2.Items.Add(text);
            }
            comboBox2.SelectedIndex = 0;

            if (ICalled)
            {
                label4.Text = _challenger.Name;
                Text = "Challenge Setup";
                btnChallenge.Text = "&Challenge!";
                btnCancel.Text = "C&ancel";
            }
            else
            {
                label4.Text = "You have been challenged!";
                Text = _challenger.Name + " - Ping Speed: " + _challenger.Ping;
                btnChallenge.Text = "&Accept";
                btnCancel.Text = "&Refuse";
                comboBox1.Visible = false;
                comboBox2.Visible = false;
                lblTerrain.Visible = false;
            }

            ShowHiddenTeamWarnings();

            listView1.Items.Clear();
            listView1.CheckBoxes = ICalled;
            if (!ICalled)
            {
                _client.ChallengePending = true;
                for (int rule = 1; rule <= ChallengeRules.RuleCount; rule++)
                {
                    if (_settings.Rules[rule])
                        AddRule(rule);
                }

                string display = GameMode is >= 1 and <= 4 ? ChallengeRules.ModeText[GameMode - 1] : "";
                display += _settings.Terrain == 0
                    ? " - Random Terrain"
                    : " - " + ChallengeRules.TerrainText((TerrainTypes)(_settings.Terrain - 1));
                lblMode.Text = display;
            }

            lblExtraInfo.Text = _challenger.Description;
            ShowRecord();

            progressBar1.Value = Math.Clamp(_challenger.TeamPower, progressBar1.Minimum, progressBar1.Maximum);
            ColorBars();

            if (_challenger.Id == me.Id)
            {
                btnChallenge.Enabled = false;
                toolTip1.SetToolTip(btnChallenge, "Can't challenge self!");
            }
            else
            {
                btnChallenge.Enabled = true;
                toolTip1.SetToolTip(btnChallenge, "");
            }
            if (CleanVersion(_challenger.Version) != CleanVersion(_client.You.Version))
            {
                btnChallenge.Enabled = false;
                toolTip1.SetToolTip(btnChallenge, "Version Conflict!");
                Text += " - Version Conflict!";
            }
            if (CleanVersion(_client.Info.ServerVersion) != CleanVersion(_client.You.Version))
            {
                btnChallenge.Enabled = false;
                toolTip1.SetToolTip(btnChallenge, "Server Version Conflict!");
                Text += " - Server Version Conflict!";
            }

            if (ICalled)
            {
                for (int rule = 1; rule <= ChallengeRules.RuleCount; rule++)
                {
                    if (ChallengeRules.RuleText(rule) != "")
                        AddRule(rule);
                }

                if (!(_challenger.StadiumOk && me.StadiumOk))
                    SetGhosted((int)RuleEnum.nbStadiumMode, true, Color.Red);

                ReadRules();
                ModeChanged();
            }
            else
            {
                FlashUntilFocused();
            }

            DoImages();
        }

        private static bool RbyCompatible(Player p) => (p.Compatibility & 1) != 0; // -- .Compatibility(0)

        private static string CleanVersion(string? version) => (version ?? "").Trim().TrimEnd('\0');

        private void AddRule(int rule)
        {
            var item = new ListViewItem(ChallengeRules.RuleText(rule))
            {
                Tag = rule,
                ToolTipText = ChallengeRules.RuleToolTip(rule)
            };
            listView1.Items.Add(item);
        }

        private ListViewItem? RuleItem(RuleEnum rule) =>
            listView1.Items.Cast<ListViewItem>().FirstOrDefault(i => (int)i.Tag! == (int)rule);

        private void SetGhosted(int rule, bool ghosted, Color? ghostColor = null)
        {
            ListViewItem? item = RuleItem((RuleEnum)rule);
            if (item == null)
                return;

            if (ghosted)
            {
                item.Checked = false; // -- Before it's ghosted, or ItemCheck would undo it.
                item.ForeColor = ghostColor ?? GhostColor;
                _ghosted.Add(rule);
            }
            else
            {
                _ghosted.Remove(rule);
                item.ForeColor = SystemColors.WindowText;
            }
        }

        private void ShowHiddenTeamWarnings()
        {
            lblUbers.Text = "";
            if (_challenger.ShowTeam)
                return;

            int ubers = 0;
            bool spec = false;
            var team = _challenger.Team ?? new List<TeamMember>();
            for (int x = 0; x < team.Count; x++)
            {
                if (team[x].Pokemon.Uber)
                    ubers++;
                if (ChallengeRules.CompatVersion(_challenger.GameVersion) == BattleModes.nbAdvBattle && team[x].Pokemon.No == 202)
                    ubers++; // -- Ugh wobb.
                for (int y = x + 1; y < team.Count; y++)
                {
                    if (team[x].Pokemon.No == team[y].Pokemon.No)
                        spec = true;
                }
            }

            string temp = ubers switch
            {
                1 => "1 Uber",
                > 1 => ubers + " Ubers",
                _ => ""
            };
            if (spec)
                temp = temp.Length == 0 ? "Species Clause" : temp + " - Spec.";
            lblUbers.Text = temp;
        }

        private void ShowRecord()
        {
            int total = _challenger.Wins + _challenger.Losses + _challenger.Ties;
            lblWins.Text = RecordText(_challenger.Wins, total);
            lblLosses.Text = RecordText(_challenger.Losses, total);
            lblTies.Text = RecordText(_challenger.Ties, total);
        }

        private static string RecordText(int count, int total) =>
            count + Environment.NewLine + (total > 0 ? (count * 100 / total) + "%" : "---%");

        private void ColorBars()
        {
            progressBar1.ForeColor = progressBar1.Value switch
            {
                <= 25 => Color.Red,
                <= 50 => Color.Yellow,
                _ => Color.Lime
            };
        }

        private void ModeChanged()
        {
            if (!ICalled)
                return;

            int x = comboBox1.SelectedIndex + 1;
            SetGhosted((int)RuleEnum.nbRandbat, !_canBattle[x]);
            SetGhosted((int)RuleEnum.nbPresentRule, x >= 3);
            DoImages();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ModeChanged();
        }

        private void listView1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var rule = (int)listView1.Items[e.Index].Tag!;
            if (_ghosted.Contains(rule))
                e.NewValue = e.CurrentValue;
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if ((int)e.Item.Tag! == (int)RuleEnum.nbRandbat)
            {
                tmrRandbat.Enabled = e.Item.Checked;
                DoImages();
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!ICalled)
                listView1.SelectedItems.Clear(); // -- The received rules are read-only.
        }

        private void SaveRules()
        {
            var rules = new List<bool>();
            for (int rule = 1; rule <= ChallengeRules.RuleCount; rule++)
                rules.Add(RuleItem((RuleEnum)rule)?.Checked ?? false);

            Configuration.CurrentSettings.BattleRules = rules;
            Configuration.CurrentSettings.BattleMode = comboBox1.SelectedIndex;
            Configuration.Save();
        }

        private void ReadRules()
        {
            List<bool> saved = Configuration.CurrentSettings.BattleRules ?? new List<bool>();
            foreach (ListViewItem item in listView1.Items)
            {
                var rule = (int)item.Tag!;
                if (!_ghosted.Contains(rule))
                    item.Checked = rule - 1 < saved.Count && saved[rule - 1];
            }

            ListViewItem? randbat = RuleItem(RuleEnum.nbRandbat);
            if (_challenger.Id == _client.Me.Id && randbat != null)
                randbat.Checked = false;

            int gmTemp = Configuration.CurrentSettings.BattleMode;
            if (gmTemp > 3 || gmTemp < 0)
                gmTemp = 2;
            if (randbat is { Checked: true })
                comboBox1.SelectedIndex = gmTemp;
            else if (ChallengeRules.CompatVersion(_client.Me.GameVersion) == BattleModes.nbAdvBattle && gmTemp == 3)
                comboBox1.SelectedIndex = 3;
        }

        private bool RandbatOn() =>
            ICalled
                ? RuleItem(RuleEnum.nbRandbat)?.Checked ?? false
                : _settings.Rules[(int)RuleEnum.nbRandbat];

        private void DoImages()
        {
            var pictureBoxes = new[] { pk1, pk2, pk3, pk4, pk5, pk6 };
            bool randBat = RandbatOn();
            if (randBat && !tmrRandbat.Enabled)
                tmrRandbat.Enabled = true;

            var team = _challenger.Team ?? new List<TeamMember>();
            for (int x = 0; x < 6; x++)
            {
                string? graphic;
                string tip;
                TeamMember? poke = x < team.Count ? team[x] : null;

                if (randBat)
                {
                    graphic = _randPoke[x + 1].ToString("000") + "_1";
                    tip = "Random";
                }
                else if (_challenger.ShowTeam || _challenger.Id == _client.Me.Id)
                {
                    graphic = poke == null ? null : "team:" + x;
                    tip = poke?.Pokemon.Name ?? "";
                }
                else
                {
                    graphic = "000rs";
                    tip = poke != null && poke.Pokemon.Uber ? "Hidden (Über)" : "Hidden";
                }

                toolTip1.SetToolTip(pictureBoxes[x], tip);
                if (Equals(pictureBoxes[x].Tag, graphic))
                    continue;

                pictureBoxes[x].Tag = graphic;
                byte[]? img = null;
                try
                {
                    if (graphic == null)
                        img = null;
                    else if (graphic.StartsWith("team:"))
                        img = GraphicsDatabase.GetSprite(poke!, GraphicsMode.nbGFXSml, false);
                    else
                        img = GraphicsDatabase.GetSprite(graphic);
                }
                catch (Exception ex)
                {
                    Logger.Log(LogType.Warning, $"Couldn't load sprite {graphic}: {ex.Message}");
                }

                pictureBoxes[x].Image = img == null ? null : Image.FromStream(new MemoryStream(img));
                pictureBoxes[x].SizeMode = graphic == "000rs" ? PictureBoxSizeMode.StretchImage : PictureBoxSizeMode.Normal;
            }
        }

        private void tmrRandbat_Tick(object sender, EventArgs e)
        {
            int z = comboBox1.SelectedIndex switch
            {
                0 => 151,
                1 => 251,
                _ => 385
            };

            // -- Re-rolls one slot per tick, sweeping 1 to 6 and back.
            _randStep++;
            if (_randStep > 12)
                _randStep = 1;
            int a = _randStep > 6 ? 13 - _randStep : _randStep;

            int pick;
            do
            {
                pick = _random.Next(1, z + 1);
            } while (pick == 201 || Enumerable.Range(1, 6).Any(y => y != a && _randPoke[y] == pick));
            _randPoke[a] = pick;

            DoImages();
        }

        private void btnChallenge_Click(object sender, EventArgs e)
        {
            if (ICalled)
            {
                var rules = new bool[ChallengeRules.RuleCount + 1];
                foreach (ListViewItem item in listView1.Items)
                    rules[(int)item.Tag!] = item.Checked;

                _client.SendChallenge(_challenger, new ChallengeSettings
                {
                    Mode = comboBox1.SelectedIndex + 1,
                    Terrain = comboBox2.SelectedIndex,
                    Rules = rules
                });
                PressedOk = true;
                btnChallenge.Enabled = false;
                SaveRules();
            }
            else
            {
                _client.AcceptChallenge(new ChallengeSettings
                {
                    Mode = GameMode,
                    Terrain = _settings.Terrain,
                    Rules = _settings.Rules
                });
                _replied = true;
                Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (ICalled)
            {
                ICalled = false;
                _client.CancelChallenge(PressedOk);
            }
            else
            {
                _client.RefuseChallenge();
            }

            _replied = true;
            Close();
        }

        private void ClientOnChallengeEnded()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new EmptyEventArgs(ClientOnChallengeEnded));
                return;
            }

            if (IsDisposed)
                return;
            btnChallenge.Enabled = true;
            _replied = true;
            Close();
        }

        private void ChallengeWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            // -- VB's Form_QueryUnload: closing it yourself counts as Cancel / Refuse.
            if (!_replied && e.CloseReason == CloseReason.UserClosing)
                btnCancel_Click(this, EventArgs.Empty);
        }

        private void ChallengeWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
            tmrRandbat.Enabled = false;
            _client.ChallengeEnded -= ClientOnChallengeEnded;
            _client.ChallengeWindowOpen = false;
            if (!_client.ChallengePending)
                _client.ChallengeNumber = 0;
            if (Current == this)
                Current = null;
        }

        #region Window flashing
        // -- VB flashed the window once a second until it was clicked; FlashWindowEx does the same on its own.
        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        private const uint FLASHW_ALL = 3;
        private const uint FLASHW_TIMERNOFG = 12;

        [DllImport("user32.dll")]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        private void FlashUntilFocused()
        {
            Form target = MdiParent ?? this;
            var info = new FLASHWINFO
            {
                cbSize = (uint)Marshal.SizeOf<FLASHWINFO>(),
                hwnd = target.Handle,
                dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG,
                uCount = uint.MaxValue,
                dwTimeout = 1000
            };
            FlashWindowEx(ref info);
        }
        #endregion
    }
}
