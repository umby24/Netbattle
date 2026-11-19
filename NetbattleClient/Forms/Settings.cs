using System;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Netbattle.Core.Database;

namespace Netbattle.Forms
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void txtRegAddr_TextChanged(object sender, EventArgs e)
        {
            UserSettings.CurrentSettings.RegistryAddress = txtRegAddr.Text;
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            txtRegAddr.Text = UserSettings.CurrentSettings.RegistryAddress;
            chkPlayerSignOn.Checked = Configuration.CurrentSettings.MessageToggles[0];
            chkPlayerSignOff.Checked = Configuration.CurrentSettings.MessageToggles[1];
            chkBattleStart.Checked = Configuration.CurrentSettings.MessageToggles[2];
            chkBattleResults.Checked = Configuration.CurrentSettings.MessageToggles[3];
            chkReturnFromBattle.Checked = Configuration.CurrentSettings.MessageToggles[4];
            chkShowAway.Checked = Configuration.CurrentSettings.MessageToggles[5];
            chkShowBack.Checked = Configuration.CurrentSettings.MessageToggles[6];
            chkShowIgnore.Checked = Configuration.CurrentSettings.MessageToggles[7];
            chkShowTeamChange.Checked = Configuration.CurrentSettings.MessageToggles[8];
            chkShowNickChange.Checked = Configuration.CurrentSettings.MessageToggles[9];
            chkLogsAutosave.Checked = Configuration.CurrentSettings.LogsAutosave;
            chkLogsSave.Checked = Configuration.CurrentSettings.LogsPrompt;
            chkReplayAutosave.Checked = Configuration.CurrentSettings.ReplayAutosave;
            chkReplays.Checked = Configuration.CurrentSettings.ReplayPrompt;
            chkEnableBackground.Checked = Configuration.CurrentSettings.EnableBackground;
            chkAnimateHP.Checked = Configuration.CurrentSettings.AnimateHpBars;
            chkSounds.Checked = Configuration.CurrentSettings.Sounds;
            chkMusic.Checked = Configuration.CurrentSettings.Music;
            chkTrainerPrefix.Checked = Configuration.CurrentSettings.PlayerNamePrefix;
            chkUseNicks.Checked = Configuration.CurrentSettings.UseNicknames;
            chkTimestamp.Checked = Configuration.CurrentSettings.UseTimestamp;
            txtTSFormat.Text = Configuration.CurrentSettings.TimestampFormat;
            chkParseUrls.Checked = Configuration.CurrentSettings.ParseUrls;
            
            if (Configuration.CurrentSettings.ShowTeam) {
                rdoShowTeam.Checked = true;
                rdoHideTeam.Checked = false;
            }
            else {
                rdoShowTeam.Checked = false;
                rdoHideTeam.Checked = true;
            }

            if (Configuration.CurrentSettings.ExtendedMessages) {
                rdoExtended.Checked = true;
                rdoNormal.Checked = false;
            }
            else {
                rdoExtended.Checked = false;
                rdoNormal.Checked = true;
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e) {
        }

        private void btnOk_Click(object sender, EventArgs e) {
            Configuration.CurrentSettings.MessageToggles =
                new[] {
                    chkPlayerSignOn.Checked, 
                    chkPlayerSignOff.Checked,
                    chkBattleStart.Checked, 
                    chkBattleResults.Checked,
                    chkReturnFromBattle.Checked,
                    chkShowAway.Checked,
                    chkShowBack.Checked,
                    chkShowIgnore.Checked,
                    chkShowTeamChange.Checked,
                    chkShowNickChange.Checked
                };
            
            Configuration.CurrentSettings.RegistryIp = txtRegAddr.Text;
            Configuration.CurrentSettings.AnimateHpBars = chkAnimateHP.Checked;
            Configuration.CurrentSettings.EnableBackground = chkEnableBackground.Checked;
            Configuration.CurrentSettings.ExtendedMessages = rdoExtended.Checked;
            Configuration.CurrentSettings.LogsAutosave = chkLogsAutosave.Checked;
            Configuration.CurrentSettings.LogsPrompt = chkLogsSave.Checked;
            Configuration.CurrentSettings.Music = chkMusic.Checked;
            Configuration.CurrentSettings.Sounds = chkSounds.Checked;
            Configuration.CurrentSettings.PlayerNamePrefix = chkTrainerPrefix.Checked;
            Configuration.CurrentSettings.ReplayAutosave = chkReplayAutosave.Checked;
            Configuration.CurrentSettings.ReplayPrompt = chkReplays.Checked;
            Configuration.CurrentSettings.ShowTeam = rdoShowTeam.Checked;
            Configuration.CurrentSettings.UseNicknames = chkUseNicks.Checked;
            Configuration.CurrentSettings.ParseUrls = chkParseUrls.Checked;
            Configuration.CurrentSettings.TimestampFormat = txtTSFormat.Text;
            Configuration.CurrentSettings.UseTimestamp = chkTimestamp.Checked;
            
            Configuration.Save();
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e) {
            Close();
        }

        private void btnStoredPassword_Click(object sender, EventArgs e) {
            var pwWind = new ServerPassword();
            pwWind.Text = "Edit Saved Password";
            pwWind.SetupSavedPw();
            var result = pwWind.ShowDialog();
            // -- Yeah yeah I know, its just how the original did it, and we're copying it. so yeah. We'll do something better later.
            var hasher = new MD5CryptoServiceProvider();
            hasher.Initialize();
            var hashBytes = hasher.ComputeHash(Encoding.ASCII.GetBytes(pwWind.Password));
            Configuration.CurrentSettings.StoredPassword = Convert.ToHexString(hashBytes);
            Configuration.Save();
        }
    }
}