using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Database;

namespace Netbattle.Forms {
    public partial class TeamBuilder : Form {
        private Pokemon _currentPokemon;
        private bool _changed;
        private UserSettings _workingSettings;

        public TeamBuilder(Form parent) {
            this.MdiParent = parent;
            InitializeComponent();
            _changed = false;
        }

        private void TeamBuilder_Load(object sender, EventArgs e) {
            dropGraphics.SelectedIndex = (int)UserSettings.CurrentSettings.CurrentGraphicsMode;

            // -- Loads the pokemon!
            var pokemonNames = PokemonDatabase.BasePokemon.Select(a => a.Name).ToArray();
            dropPokemon.Items.AddRange(pokemonNames);

            // -- Load your user info..
            txtLoseMessage.Text = UserSettings.CurrentSettings.LoseMessage;
            txtWinMessage.Text = UserSettings.CurrentSettings.WinMessage;
            txtUsername.Text = UserSettings.CurrentSettings.Username;
            txtExtraInfo.Text = UserSettings.CurrentSettings.MoreInfo;
            listView1.Items[UserSettings.CurrentSettings.IconUsed].Selected = true;

            if (UserSettings.CurrentSettings.Team[0] != null) {
                dropPokemon.SelectedIndex = UserSettings.CurrentSettings.Team[0].No - 1;
                txtNickname.Text = UserSettings.CurrentSettings.Team[0].Nickname;
            }

            _workingSettings = UserSettings.CurrentSettings;
        }

        private void dropPokemon_SelectedIndexChanged(object sender, EventArgs e) {
            var pkmnObj = PokemonDatabase.BasePokemon.FirstOrDefault(a => a.Name == (string) dropPokemon.SelectedItem);
            _currentPokemon = pkmnObj;

            var img = GraphicsDatabase.GetSprite(pkmnObj, UserSettings.CurrentSettings.CurrentGraphicsMode, false);
            using (var ms = new MemoryStream(img)) {
                picImage.Image = Image.FromStream(ms);
            }

            // -- TODO: How will this get populated..??
            txtNickname.Text = pkmnObj.Nickname;
            lblHp.Text = "HP: " + pkmnObj.HP;
            lblAttack.Text = "Attack: " + pkmnObj.Attack;
            lblDefense.Text = "Defense: " + pkmnObj.Defense;
            lblSpeed.Text = "Speed: " + pkmnObj.Speed;
            lblSpecialAttack.Text = "Sp. Attack: " + pkmnObj.SpecialAttack;
            lblSpecialDefense.Text = "Sp. Defense: " + pkmnObj.SpecialDefense;

            // -- Format No to '000' + e for emerald.
            PopulateMoves();
        }

        
        private void PopulateMoves() {
            if (_currentPokemon == null || _currentPokemon.No == 0) {
                listMoves.Items.Clear();
                return;
            }

            var moves = _currentPokemon.GetAdvMoves();
            listMoves.Items.Clear();

            foreach (Move move in moves) {
                
                var item = new ListViewItem();
                item.Text = move.Name;
                var subItem2 = new ListViewItem.ListViewSubItem(item, move.Power.ToString());
                var subItem3 = new ListViewItem.ListViewSubItem(item, move.Accuracy.ToString() + "%");
                var subItem4 = new ListViewItem.ListViewSubItem(item, move.PP.ToString());
                var subItem5 = new ListViewItem.ListViewSubItem(item, move.Source.ToString());
                item.SubItems.Add(subItem2);
                item.SubItems.Add(subItem3);
                item.SubItems.Add(subItem4);
                item.SubItems.Add(subItem5);

                listMoves.Items.Add(item);
            }
        }

        private void dropGraphics_SelectedIndexChanged(object sender, EventArgs e) {

        }

        private void TeamBuilder_FormClosing(object sender, FormClosingEventArgs e) {
            DialogResult mb = MessageBox.Show("You have unsaved changes. Would you like to save now?", "Unsaved Changes",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

            if (mb == DialogResult.Cancel) {
                e.Cancel = true;
                return;
            }

            if (mb == DialogResult.Yes) {
                UpdateWorking();
                UserSettings.Save();
            }
        }

        private void UpdateWorking() {
            UserSettings.CurrentSettings.CurrentGraphicsMode = (GraphicsMode) dropGraphics.SelectedIndex;
            UserSettings.CurrentSettings.LoseMessage = txtLoseMessage.Text;
            UserSettings.CurrentSettings.WinMessage = txtWinMessage.Text;
            UserSettings.CurrentSettings.MoreInfo = txtExtraInfo.Text;
            UserSettings.CurrentSettings.IconUsed = (byte) listView1.SelectedIndices[0];
            UserSettings.CurrentSettings.Username = txtUsername.Text;

            if (dropPokemon.SelectedIndex != -1) {
                UserSettings.CurrentSettings.Team[0] = PokemonDatabase.BasePokemon[dropPokemon.SelectedIndex];
                var test = UserSettings.CurrentSettings.Team[0].ToString();
                Console.WriteLine("a");
            }

        }
    }
}
