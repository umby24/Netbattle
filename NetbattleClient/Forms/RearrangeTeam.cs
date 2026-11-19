using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.VisualBasic.ApplicationServices;
using Netbattle.Core.Common;
using Netbattle.Core.Database;

namespace Netbattle.Forms
{
    // -- Layout converted from Rearrange.frm. None of its code has been ported yet.
    public partial class RearrangeTeam : Form
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public bool DidChange { get; set; }

        private TeamMember[] _origTeam;
        
        public RearrangeTeam()
        {
            InitializeComponent();
            DidChange = false;
            RefreshBox();
            _origTeam = new TeamMember[6];
            UserSettings.CurrentSettings.Team.CopyTo(_origTeam, 0);
        }

        private void MoveUp_Click(object sender, EventArgs e) {
            if (List1.SelectedItem == null)
                return;

            if (List1.SelectedIndex == 0)
                return;

            var tmpPkmn = UserSettings.CurrentSettings.Team[List1.SelectedIndex - 1];
            UserSettings.CurrentSettings.Team[List1.SelectedIndex - 1] =
                UserSettings.CurrentSettings.Team[List1.SelectedIndex];
            UserSettings.CurrentSettings.Team[List1.SelectedIndex] = tmpPkmn;
            DidChange = true;
            RefreshBox();
        }

        private void MoveDown_Click(object sender, EventArgs e) {
            if (List1.SelectedItem == null)
                return;

            if (List1.SelectedIndex == 5)
                return;

            var tmpPkmn = UserSettings.CurrentSettings.Team[List1.SelectedIndex + 1];
            UserSettings.CurrentSettings.Team[List1.SelectedIndex + 1] =
                UserSettings.CurrentSettings.Team[List1.SelectedIndex];
            UserSettings.CurrentSettings.Team[List1.SelectedIndex] = tmpPkmn;
            DidChange = true;
            RefreshBox();
        }

        private void CancelButton__Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            UserSettings.CurrentSettings.Team = _origTeam;
        }

        private void OKButton_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.OK;
        }

        private void RefreshBox() {
            List1.Items.Clear();
            foreach (TeamMember member in UserSettings.CurrentSettings.Team) {
                List1.Items.Add($"{member.Nickname} ({member.Pokemon.Name})");
            }
        }
    }
}
