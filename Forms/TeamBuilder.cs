using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Database;

namespace Netbattle.Forms {
    public partial class TeamBuilder : Form {
        private Pokemon _currentPokemon;
        private List<Move> _currentMoveset;
        private bool _changed;
        private UserSettings _workingSettings;
        private int _currentSlot = -1;
        private bool _switching = false;

        public TeamBuilder(Form parent) {
            this.MdiParent = parent;
            InitializeComponent();
            _changed = false;
        }

        private void TeamBuilder_Load(object sender, EventArgs e) {
            dropGraphics.SelectedIndex = (int)UserSettings.CurrentSettings.CurrentGraphicsMode;

            // -- Loads the pokemon!
            string[] pokemonNames = PokemonDatabase.BasePokemon.Select(a => a.Name).ToArray();
            dropPokemon.Items.AddRange(pokemonNames);

            // -- Load your user info..
            txtLoseMessage.Text = UserSettings.CurrentSettings.LoseMessage;
            txtWinMessage.Text = UserSettings.CurrentSettings.WinMessage;
            txtUsername.Text = UserSettings.CurrentSettings.Username;
            txtExtraInfo.Text = UserSettings.CurrentSettings.MoreInfo;
            listView1.Items[UserSettings.CurrentSettings.IconUsed].Selected = true;

            _workingSettings = UserSettings.CurrentSettings;

            if (UserSettings.CurrentSettings.Team[0] != null) {
                LoadPokemonSlot(0);
            }


            for (var i = 0; i < 6; i++) {
                if (_workingSettings.Team[i] != null) {
                    // -- TODO: Set button icon to red.
                }
            }
        }

        private void dropPokemon_SelectedIndexChanged(object sender, EventArgs e) {
            if (dropPokemon.SelectedItem == null || _switching)
                return;

            Pokemon pkmnObj =
                PokemonDatabase.BasePokemon.FirstOrDefault(a => a.Name == (string)dropPokemon.SelectedItem);
            _currentPokemon = pkmnObj;
            _workingSettings.Team[_currentSlot] = _currentPokemon;

            // -- TODO: Dry?
            txtNickname.Text = pkmnObj.Nickname;
            lblHp.Text = "HP: " + pkmnObj.HP;
            lblAttack.Text = "Attack: " + pkmnObj.Attack;
            lblDefense.Text = "Defense: " + pkmnObj.Defense;
            lblSpeed.Text = "Speed: " + pkmnObj.Speed;
            lblSpecialAttack.Text = "Sp. Attack: " + pkmnObj.SpecialAttack;
            lblSpecialDefense.Text = "Sp. Defense: " + pkmnObj.SpecialDefense;

            LoadPokemonGraphic();

            // -- Format No to '000' + e for emerald.
            PopulateMoves();
        }


        private void PopulateMoves() {
            if (_currentPokemon == null || _currentPokemon.No == 0) {
                listMoves.Items.Clear();
                return;
            }

            _currentMoveset = _currentPokemon.GetAdvMoves().ToList();
            listMoves.Items.Clear();

            foreach (Move move in _currentMoveset) {

                var item = new ListViewItem { Text = move.Name };
                var subItem2 = new ListViewItem.ListViewSubItem(item, move.Power.ToString());
                var subItem3 = new ListViewItem.ListViewSubItem(item, move.Accuracy + "%");
                var subItem4 = new ListViewItem.ListViewSubItem(item, move.PP.ToString());
                var subItem5 = new ListViewItem.ListViewSubItem(item, move.Source);
                item.SubItems.Add(subItem2);
                item.SubItems.Add(subItem3);
                item.SubItems.Add(subItem4);
                item.SubItems.Add(subItem5);

                listMoves.Items.Add(item);
            }
        }

        private void dropGraphics_SelectedIndexChanged(object sender, EventArgs e) {
            if (_workingSettings != null) {
                _workingSettings.CurrentGraphicsMode = (GraphicsMode)dropGraphics.SelectedIndex;
                LoadPokemonGraphic();
            }
        }

        private void TeamBuilder_FormClosing(object sender, FormClosingEventArgs e) {
            DialogResult mb = MessageBox.Show("You have unsaved changes. Would you like to save now?",
                "Unsaved Changes",
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
            UserSettings.CurrentSettings.CurrentGraphicsMode = (GraphicsMode)dropGraphics.SelectedIndex;
            UserSettings.CurrentSettings.LoseMessage = txtLoseMessage.Text;
            UserSettings.CurrentSettings.WinMessage = txtWinMessage.Text;
            UserSettings.CurrentSettings.MoreInfo = txtExtraInfo.Text;
            UserSettings.CurrentSettings.IconUsed = (byte)listView1.SelectedIndices[0];
            UserSettings.CurrentSettings.Username = txtUsername.Text;
            UserSettings.CurrentSettings.Team = _workingSettings.Team;
        }

        private void SwitchPokemonButtonClick(object sender, EventArgs e) {
            btnPoke1.FlatStyle = FlatStyle.Standard;
            btnPoke2.FlatStyle = FlatStyle.Standard;
            btnPoke3.FlatStyle = FlatStyle.Standard;
            btnPoke4.FlatStyle = FlatStyle.Standard;
            btnPoke5.FlatStyle = FlatStyle.Standard;
            btnPoke6.FlatStyle = FlatStyle.Standard;

            var btnObj = (Button)sender;
            LoadPokemonSlot(int.Parse(btnObj.Text.Replace("PKMN ", "")) - 1);
            btnObj.FlatStyle = FlatStyle.Popup;
        }

        private void LoadPokemonGraphic() {
            byte[] img = GraphicsDatabase.GetSprite(_currentPokemon, _workingSettings.CurrentGraphicsMode, false);

            using (var ms = new MemoryStream(img)) {
                picImage.Image = Image.FromStream(ms);
            }
        }

        private void LoadPokemonSlot(int slot) {
            _switching = true;
            // -- Save working pokemon..
            if (_currentSlot != -1) {
                _workingSettings.Team[_currentSlot] = _currentPokemon;
            }

            _currentSlot = slot;
            _currentPokemon = null;

            if (_workingSettings.Team[slot] == null) {
                dropPokemon.SelectedIndex = -1;
                txtNickname.Text = "";
                picImage.Image = null;
                PopulateMoves();
                _switching = false;
                return;
            }

            dropPokemon.SelectedIndex = _workingSettings.Team[slot].No - 1;
            txtNickname.Text = _workingSettings.Team[slot].Nickname;
            _currentPokemon = _workingSettings.Team[slot];
            dropItem.SelectedIndex = (int) _currentPokemon.Item;

            LoadPokemonGraphic();

            lblHp.Text = "HP: " + _currentPokemon.HP;
            lblAttack.Text = "Attack: " + _currentPokemon.Attack;
            lblDefense.Text = "Defense: " + _currentPokemon.Defense;
            lblSpeed.Text = "Speed: " + _currentPokemon.Speed;
            lblSpecialAttack.Text = "Sp. Attack: " + _currentPokemon.SpecialAttack;
            lblSpecialDefense.Text = "Sp. Defense: " + _currentPokemon.SpecialDefense;
            PopulateMoves();

            for (var i = 0; i < 4; i++) {
                if (_currentPokemon.Move[i] != -1)
                    SelectMove(MoveDatabase.Moves[_currentPokemon.Move[i]].Name);
            }

            UpdateMoveBoxes();
            _switching = false;
        }

        private int GetMoveIndex(Move move) {
            for (var i = 0; i < 4; i++) {
                if (_currentPokemon.Move[i] == move.ID && move.ID != -1) {
                    return i;
                }
            }

            return -1;
        }

        private void SelectMove(string name) {
            foreach (ListViewItem item in listMoves.Items) {
                if (item.Text == name) {
                    item.Checked = true;
                    break;
                }
            }
        }
        private void listMoves_ItemChecked(object sender, ItemCheckedEventArgs e) {
            if (e.Item.Index >= _currentMoveset.Count || _switching)
                return;

            Move moveObj = _currentMoveset.FirstOrDefault(a => a.Name == e.Item.Text);

            if (e.Item.Checked == false) {
                // -- Remove this move..
                int moveIndex = GetMoveIndex(moveObj);
                if (moveIndex == -1) return;

                _currentPokemon.Move[moveIndex] = -1;
                UpdateMoveBoxes();
                return;
            }

            if (_currentPokemon.Move[0] != -1 && _currentPokemon.Move[1] != -1 && _currentPokemon.Move[2] != -1 &&
                _currentPokemon.Move[3] != -1) {
                MessageBox.Show("A pokemon cannot have more than 4 moves!", "Too many moves!", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                e.Item.Checked = false;
                return;
            }

            for (var i = 0; i < 4; i++) {
                if (_currentPokemon.Move[i] == -1) {
                    _currentPokemon.Move[i] = moveObj.ID;
                    break;
                }
            }

            UpdateMoveBoxes();
        }

        private void UpdateMoveBoxes() {
            txtMove1.Text = "";
            txtMove2.Text = "";
            txtMove3.Text = "";
            txtMove4.Text = "";

            Move move1 = null, move2 = null, move3 = null, move4 = null;
            if (_currentPokemon.Move[0] != -1)
                move1 = MoveDatabase.Moves[_currentPokemon.Move[0]];
            if (_currentPokemon.Move[1] != -1)
                move2 = MoveDatabase.Moves[_currentPokemon.Move[1]];
            if (_currentPokemon.Move[2] != -1)
                move3 = MoveDatabase.Moves[_currentPokemon.Move[2]];
            if (_currentPokemon.Move[3] != -1)
                move4 = MoveDatabase.Moves[_currentPokemon.Move[3]];

            if (move1 != null)
                txtMove1.Text = move1.Name;

            if (move2 != null)
                txtMove2.Text = move2.Name;

            if (move3 != null)
                txtMove3.Text = move3.Name;

            if (move4 != null)
                txtMove4.Text = move4.Name;
        }

        private void txtNickname_TextChanged(object sender, EventArgs e) {
            if (_currentPokemon != null && txtNickname.Text != null && txtNickname.Text != _currentPokemon.Nickname)
                _currentPokemon.Nickname = txtNickname.Text;
        }

        private void dropItem_SelectedIndexChanged(object sender, EventArgs e) {
            if (_currentPokemon != null)
            _currentPokemon.Item = (Items)Enum.Parse(typeof(Items), "nb" + dropItem.Text.Replace(" ", ""));
        }
    }
}