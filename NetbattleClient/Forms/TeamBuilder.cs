using Netbattle.Common;
using Netbattle.Core.Battle;
using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;
using Netbattle.Core.Files;

namespace Netbattle.Forms {
    public partial class TeamBuilder : Form {
        private TeamMember _currentPokemon;
        private List<Move> _currentMoveset;
        private bool _changed;
        private UserSettings _workingSettings;
        private int _currentSlot = 0;
        private bool _switching = false;
        private ServerWindow _serverWindow;

        public TeamBuilder(Form parent) {
            this.MdiParent = parent;
            InitializeComponent();
            _changed = false;
            if (parent is ServerWindow window) {
                _serverWindow = window;
                // -- Send Away Command! TODO:;
            }
            if (GraphicsDatabase.HasColGraphics) {
                dropGraphics.Items.Add("Colosseum");
            }
        }

        private void TeamBuilder_Load(object sender, EventArgs e) {
            if (_serverWindow != null) {
                if (!_serverWindow.awayToolStripMenuItem.Checked) {
                    _serverWindow.ToggleAway();
                }
            }

            dropGraphics.SelectedIndex = (int)UserSettings.CurrentSettings.CurrentGraphicsMode-1;

            // -- Loads the pokemon!
            string[] pokemonNames = PokemonDatabase.BasePokemon.Select(a => a.Name).ToArray();
            dropPokemon.Items.AddRange(pokemonNames);

            // -- Load your user info..
            txtLoseMessage.Text = UserSettings.CurrentSettings.LoseMessage;
            txtWinMessage.Text = UserSettings.CurrentSettings.WinMessage;
            txtUsername.Text = UserSettings.CurrentSettings.Username;
            txtExtraInfo.Text = UserSettings.CurrentSettings.MoreInfo;
            listView1.Items[UserSettings.CurrentSettings.IconUsed-1].Selected = true;

            _workingSettings = UserSettings.CurrentSettings;
            // -- Load your team

            for (var i = 0; i < 6; i++) {
                if (UserSettings.CurrentSettings.Team[i] != null) {
                    LoadPokemonSlot(i);
                }

                if (_workingSettings.Team[i] != null) {
                    // -- If you have a pokemon in a slot, turn the icon red from gray.
                    btnPoke1.ImageIndex = 16;
                }
                else {
                    btnPoke1.ImageIndex = 15;
                }
            }

            // -- Updates menu and dropdown selections across the form
            UpdateSelections();
            // -- CenterWindow()
            if (!string.IsNullOrEmpty(_workingSettings.DbModName)) {
                // -- Set menu strip item text to Mod: [No Name] or Mod: name
            }
            else {
                // -- Set menu strip item text to 'No Mod Loaded'
            }

            // -- DoRBY()
        }

        private void RefreshCompatTree() {
        }

        private void dropPokemon_SelectedIndexChanged(object sender, EventArgs e) {
            if (dropPokemon.SelectedItem == null || _switching)
                return;

            var newMember = new TeamMember(PokemonDatabase.BasePokemon.FirstOrDefault(a => a.Name == dropPokemon.SelectedItem.ToString()));
            newMember.ApplyDefaultStats(_workingSettings.CurrentCompatibilityMode);
            newMember.Initialize(_workingSettings.CurrentCompatibilityMode);

            // -- Deoxys form check
            if (newMember.Pokemon.No == 386 || newMember.Pokemon.No == 387 || newMember.Pokemon.No == 388 || newMember.Pokemon.No == 389) {
                for (int i = 0; i < 6; i++) {
                    if ((_workingSettings.Team[i]?.Pokemon?.No == 386 || _workingSettings.Team[i]?.Pokemon?.No == 387 ||
                         _workingSettings.Team[i]?.Pokemon?.No == 388 || _workingSettings.Team[i]?.Pokemon?.No == 389) &&
                        i != _currentSlot) {
                        MessageBox.Show("You already have a " + newMember.Pokemon.Name + " on your team!", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            // -- Compatibility Check

            switch (_workingSettings.CurrentCompatibilityMode)
            {
                case CompatModes.nbTrueRBY:
                case CompatModes.nbRBYTrade:
                    if (!newMember.Pokemon.ExistRBY) {
                        MessageBox.Show("This pokemon does not exist in RBY!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    break;
                case CompatModes.nbGSCTrade:
                    if (newMember.Pokemon.No > 251)
                    {
                        MessageBox.Show("This pokemon does not exist in Gen 2!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    break;
                case CompatModes.nbTrueRuSa:
                    if (!newMember.Pokemon.ExistADV)
                    {
                        MessageBox.Show("This pokemon does not exist in RuSa!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    break;
                case CompatModes.nbTrueGSC:
                    if (!newMember.Pokemon.ExistGSC)
                    {
                        MessageBox.Show("This pokemon does not exist in GSC!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    break;
            }

            _currentPokemon = newMember;
            _workingSettings.Team[_currentSlot] = _currentPokemon;
            _changed = true;

            UpdateStatLabels();
            LoadPokemonGraphic();

            // -- Format No to '000' + e for emerald.
            PopulateMoves();
        }


        private void PopulateMoves() {
            if (_currentPokemon == null || _currentPokemon.Pokemon.No == 0) {
                listMoves.Items.Clear();
                return;
            }

            _currentMoveset = _currentPokemon.Pokemon.GetCurrentMoveset(_workingSettings.CurrentCompatibilityMode);
            listMoves.Items.Clear();

            foreach (Move move in _currentMoveset) {
                var item = new ListViewItem { Text = move.Name };
                var subItem2 = new ListViewItem.ListViewSubItem(item, move.Power.ToString());
                var subItem3 = new ListViewItem.ListViewSubItem(item, move.Accuracy + "%");
                var subItem4 = new ListViewItem.ListViewSubItem(item, move.PP.ToString());
                var subItem5 = new ListViewItem.ListViewSubItem(item, move.Source.ToFriendlyString());
                item.SubItems.Add(subItem2);
                item.SubItems.Add(subItem3);
                item.SubItems.Add(subItem4);
                item.SubItems.Add(subItem5);

                switch (move.Type) {
                    case Elements.nbNormal:
                        item.ImageIndex = 11; //0;
                        break;
                    case Elements.nbWater:
                        item.ImageIndex = 16;
                        break;
                    case Elements.nbFire:
                        item.ImageIndex = 5;
                        break;
                    case Elements.nbIce:
                        item.ImageIndex = 10; //3;
                        break;
                    case Elements.nbElectr:
                        item.ImageIndex = 3;
                        break;

                    case Elements.nbBug:
                        item.ImageIndex = 0; //5;
                        break;
                    case Elements.nbFlying:
                        item.ImageIndex = 6;
                        break;
                    case Elements.nbDragon:
                        item.ImageIndex = 2;
                        break;
                    case Elements.nbFight:
                        item.ImageIndex = 4;
                        break;
                    case Elements.nbGhost:
                        item.ImageIndex = 7;
                        break;
                    case Elements.nbGrass:
                        item.ImageIndex = 8;
                        break;
                    case Elements.nbGround:
                        item.ImageIndex = 9;
                        break;
                    case Elements.nbPoison: // -- Good!
                        item.ImageIndex = 12;
                        break;
                    case Elements.nbPsychc: // -- Good!
                        item.ImageIndex = 13;
                        break;
                    case Elements.nbRock:
                        item.ImageIndex = 14;
                        break;
                    case Elements.nbDark:
                        item.ImageIndex = 1;
                        break;
                    case Elements.nbSteel:
                        item.ImageIndex = 15;
                        break;
                    default:
                        item.ImageIndex = 11;
                        break;
                }


                listMoves.Items.Add(item);
            }
        }

        private void dropGraphics_SelectedIndexChanged(object sender, EventArgs e) {
            if (_workingSettings != null) {
                _workingSettings.CurrentGraphicsMode = (GraphicsMode)dropGraphics.SelectedIndex + 1;
                LoadPokemonGraphic();
                _changed = true;
            }
        }

        private void TeamBuilder_FormClosing(object sender, FormClosingEventArgs e) {
            foreach (var pkmn in _workingSettings.Team) {
                if (string.IsNullOrEmpty(pkmn.Nickname))
                    pkmn.Nickname = pkmn.Pokemon.Name;
            }
            
            if (!_changed) return;

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
                Configuration.CurrentSettings.LastJnb = "team.jnb";
                Configuration.Save();
            }
        }

        private void UpdateSelections() {
            fullAdvancedToolStripMenuItem.Checked = false;
            ruSaOnlyToolStripMenuItem.Checked = false;
            gSCWithTradesToolStripMenuItem.Checked = false;
            trueGSCToolStripMenuItem.Checked = false;
            rBYWithTradesToolStripMenuItem.Checked = false;
            trueRBYToolStripMenuItem.Checked = false;

            // -- Load version..
            switch (_workingSettings.CurrentCompatibilityMode) {
                case CompatModes.nbTrueRBY:
                    trueRBYToolStripMenuItem.Checked = true;
                    break;
                case CompatModes.nbRBYTrade:
                    rBYWithTradesToolStripMenuItem.Checked = true;
                    break;
                case CompatModes.nbTrueGSC:
                    trueGSCToolStripMenuItem.Checked = true;
                    break;
                case CompatModes.nbGSCTrade:
                    gSCWithTradesToolStripMenuItem.Checked = true;
                    break;
                case CompatModes.nbTrueRuSa:
                    ruSaOnlyToolStripMenuItem.Checked = true;
                    break;
                case CompatModes.nbFullAdvance:
                    fullAdvancedToolStripMenuItem.Checked = true;
                    break;
                //case CompatModes.nbModAdv:
            }

            UpdateTeamEvs();

            // -- DoRby
            // -- DoPower
            this.Text = "Teambuilder - " + NbMethods.GetTeamRank(_workingSettings.Team) + "%";
        }

        private void UpdateTeamEvs() {
            foreach (var poke in _workingSettings.Team) {
                if (poke == null) continue;
                // -- Why?
                switch (_workingSettings.CurrentCompatibilityMode) {
                    case CompatModes.nbTrueRuSa:
                    case CompatModes.nbFullAdvance:
                        if (poke.GameVersion != CompatModes.nbTrueRuSa &&
                            poke.GameVersion != CompatModes.nbFullAdvance) {
                            poke.DVs = new StatSet() {
                                [StatTypes.nbHP] = 31,
                                [StatTypes.nbAttack] = 31,
                                [StatTypes.nbDefense] = 31,
                                [StatTypes.nbSpeed] = 31,
                                [StatTypes.nbSpAttack] = 31,
                                [StatTypes.nbSpDefense] = 31
                            };
                            poke.EVs = new StatSet() {
                                [StatTypes.nbHP] = 85,
                                [StatTypes.nbAttack] = 85,
                                [StatTypes.nbDefense] = 85,
                                [StatTypes.nbSpeed] = 85,
                                [StatTypes.nbSpAttack] = 85,
                                [StatTypes.nbSpDefense] = 85
                            };
                            poke.NatureNum = 0;
                            poke.AttNum = 0;
                        }

                        break;
                    default:
                        poke.DVs = new StatSet() {
                            [StatTypes.nbHP] = 15,
                            [StatTypes.nbAttack] = 15,
                            [StatTypes.nbDefense] = 15,
                            [StatTypes.nbSpeed] = 15,
                            [StatTypes.nbSpAttack] = 15,
                            [StatTypes.nbSpDefense] = 0
                        };
                        poke.EVs = new StatSet() {
                            [StatTypes.nbHP] = 0,
                            [StatTypes.nbAttack] = 0,
                            [StatTypes.nbDefense] = 0,
                            [StatTypes.nbSpeed] = 0,
                            [StatTypes.nbSpAttack] = 0,
                            [StatTypes.nbSpDefense] = 0
                        };

                        break;
                }
            }
        }

        private void UpdateWorking() {
            UserSettings.CurrentSettings.CurrentGraphicsMode = (GraphicsMode)dropGraphics.SelectedIndex + 1;
            UserSettings.CurrentSettings.LoseMessage = txtLoseMessage.Text;
            UserSettings.CurrentSettings.WinMessage = txtWinMessage.Text;
            UserSettings.CurrentSettings.MoreInfo = txtExtraInfo.Text;
            UserSettings.CurrentSettings.IconUsed = (byte)(listView1.SelectedIndices[0] + 1);
            UserSettings.CurrentSettings.Username = txtUsername.Text;
            UserSettings.CurrentSettings.Team = _workingSettings.Team;
            UserSettings.CurrentSettings.CurrentCompatibilityMode =
                (CompatModes)_workingSettings.CurrentCompatibilityMode;
        }

        private void SwitchPokemonButtonClick(object sender, EventArgs e) {
            var btnObj = (Button)sender;
            int theSlot = int.Parse(btnObj.Text.Replace("PKMN ", "")) - 1;
            _workingSettings.Team[_currentSlot] = _currentPokemon;
            _currentSlot = theSlot;
            _currentPokemon = null;

            LoadPokemonSlot(theSlot);
            btnObj.FlatStyle = FlatStyle.Popup;
            Button[] myArr = new[] { btnPoke1, btnPoke2, btnPoke3, btnPoke4, btnPoke5, btnPoke6 };

            for (var i = 0; i < 6; i++) {
                if (_workingSettings.Team[i] != null) {
                    myArr[i].ImageIndex = 16;
                }
                else {
                    myArr[i].ImageIndex = 15;
                }
            }
        }

        private void LoadPokemonGraphic() {
            byte[] img = GraphicsDatabase.GetSprite(_currentPokemon, _workingSettings.CurrentGraphicsMode, false);

            using (var ms = new MemoryStream(img)) {
                picImage.Image = Image.FromStream(ms);
                picImage.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }

        private void LoadPokemonSlot(int slot) {
            _switching = true;
            // -- Save working pokemon..

            if (_workingSettings.Team[slot] == null) {
                dropPokemon.SelectedIndex = -1;
                txtNickname.Text = "";
                picImage.Image = null;
                PopulateMoves();
                _switching = false;
                return;
            }

            dropPokemon.SelectedIndex = _workingSettings.Team[slot].Pokemon.No - 1;
            txtNickname.Text = _workingSettings.Team[slot].Nickname;
            _currentPokemon = _workingSettings.Team[slot];
            dropItem.SelectedIndex = (int)_currentPokemon.Item;

            LoadPokemonGraphic();

            UpdateStatLabels();
            PopulateMoves();

            for (var i = 0; i < 4; i++) {
                if (_currentPokemon.Move[i] != -1 && _currentPokemon.Move[i] != 0)
                    SelectMove(MoveDatabase.Moves[_currentPokemon.Move[i]].Name);
            }

            UpdateMoveBoxes();
            _switching = false;
        }

        private void UpdateStatLabels() {
            lblHp.Text = "HP: " + _currentPokemon.CurrentStats[StatTypes.nbHP];
            lblAttack.Text = "Attack: " + _currentPokemon.CurrentStats[StatTypes.nbAttack];
            lblDefense.Text = "Defense: " + _currentPokemon.CurrentStats[StatTypes.nbDefense];
            lblSpeed.Text = "Speed: " + _currentPokemon.CurrentStats[StatTypes.nbSpeed];
            lblSpecialAttack.Text = "Sp. Attack: " + _currentPokemon.CurrentStats[StatTypes.nbSpAttack];
            lblSpecialDefense.Text = "Sp. Defense: " + _currentPokemon.CurrentStats[StatTypes.nbSpDefense];
            lblTypes.Text = "Type(s): " + _currentPokemon.Pokemon.Type1.ToFriendlyString() + " " +
                            _currentPokemon.Pokemon.Type2.ToFriendlyString().Replace("None", "");
            if (_workingSettings.CurrentCompatibilityMode == CompatModes.nbTrueRBY ||
                _workingSettings.CurrentCompatibilityMode == CompatModes.nbRBYTrade) {
                lblLevel.Text = "Lv: " + _currentPokemon.Level;
            }
            else {
                string genderString = "";
                switch (_currentPokemon.Gender) {
                    case 0:
                        genderString = "Male";
                        break;
                    case 1:
                        genderString = "Female";
                        break;
                    case 2:
                        genderString = "Genderless";
                        break;
                }
                lblLevel.Text = genderString + " Lv: " + _currentPokemon.Level;
            }
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

            if (_currentPokemon.Move[0] != 0 && _currentPokemon.Move[1] != 0 && _currentPokemon.Move[2] != 0 &&
                _currentPokemon.Move[3] != 0 && _currentPokemon.Move[0] != -1 && _currentPokemon.Move[1] != -1 && _currentPokemon.Move[2] != -1 &&
                _currentPokemon.Move[3] != -1) {
                MessageBox.Show("A pokemon cannot have more than 4 moves!", "Too many moves!", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                e.Item.Checked = false;
                return;
            }

            for (var i = 0; i < 4; i++) {
                if (_currentPokemon.Move[i] == -1 || _currentPokemon.Move[i] == 0) {
                    _currentPokemon.Move[i] = moveObj.ID;
                    break;
                }
            }

            UpdateMoveBoxes();
            // -- If you have 4 moves, set tab icon to red. Otherwise set as grey.
        }

        private void UpdateMoveBoxes() {
            txtMove1.Text = "";
            txtMove2.Text = "";
            txtMove3.Text = "";
            txtMove4.Text = "";

            Move move1 = null, move2 = null, move3 = null, move4 = null;
            if (_currentPokemon.Move[0] != -1 && _currentPokemon.Move[0] != 0)
                move1 = MoveDatabase.Moves[_currentPokemon.Move[0]];
            if (_currentPokemon.Move[1] != -1 && _currentPokemon.Move[1] != 0)
                move2 = MoveDatabase.Moves[_currentPokemon.Move[1]];
            if (_currentPokemon.Move[2] != -1 && _currentPokemon.Move[2] != 0)
                move3 = MoveDatabase.Moves[_currentPokemon.Move[2]];
            if (_currentPokemon.Move[3] != -1 && _currentPokemon.Move[3] != 0)
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

        private void DoVersionChange(CompatModes newCompatMode) {
            string changeString = "";
            
            for (var i = 0; i < _workingSettings.Team.Length; i++) {
                var p = _workingSettings.Team[i];
                if (p == null) continue;
                switch (newCompatMode) {
                    case CompatModes.nbTrueRBY:
                    case CompatModes.nbRBYTrade:
                        if (!p.Pokemon.ExistRBY) {
                            changeString += p.Nickname + " - Removed\n";
                            _workingSettings.Team[i] = null;
                        }
                        // -- Stats if they still exist
                        if (_workingSettings.CurrentCompatibilityMode == CompatModes.nbTrueRuSa || 
                            _workingSettings.CurrentCompatibilityMode == CompatModes.nbFullAdvance ||
                            _workingSettings.CurrentCompatibilityMode == CompatModes.nbModAdv) {
                            // -- Downgrade stats from GSC to RBY
                            p?.DowngradeToRBYStats();
                        }
                        // -- Items
                        if (p.Item != Items.nbNoItem) {
                            changeString += p.Nickname + " - Item removed\n";
                            p.Item = Items.nbNoItem;
                        }

                        for (var m = 0; m < 4; m++) {
                            if (p.Move[m] > 0) {
                                var mvObj = MoveDatabase.Moves[p.Move[m]];
                                if (!mvObj.RBYMove) {
                                    changeString += p.Nickname + " - Removed " + mvObj.Name + " (Newer Move)\n";
                                    p.Move[m] = -1;
                                }
                                var rules = new RBYRules();
                                
                                if (!rules.CheckMove(p.Pokemon, mvObj)) {
                                    changeString += p.Nickname + "- Removed " + mvObj.Name + " (Not Learnable in RBY)\n";
                                    p.Move[m] = -1;
                                }
                            }
                        }
                        break;
                    case CompatModes.nbTrueGSC:
                    case CompatModes.nbGSCTrade:
                        if (!p.Pokemon.ExistGSC) {
                            changeString += p.Nickname + " - Removed\n";
                            _workingSettings.Team[i] = null;
                        }
                        if (_workingSettings.CurrentCompatibilityMode == CompatModes.nbTrueRuSa || 
                            _workingSettings.CurrentCompatibilityMode == CompatModes.nbFullAdvance ||
                            _workingSettings.CurrentCompatibilityMode == CompatModes.nbModAdv) {
                            // -- Downgrade stats from GSC to RBY
                            p.DowngradeToRBYStats();
                        }
                        // -- Items
                        if (p.Item > (Items)41) {
                            changeString += p.Nickname + " - Advance-Only Item removed\n";
                            p.Item = Items.nbNoItem;
                        }
                        for (var m = 0; m < 4; m++) {
                            if (p.Move[m] > 0) {
                                var mvObj = MoveDatabase.Moves[p.Move[m]];
                                if (!mvObj.GSCMove) {
                                    changeString += p.Nickname + " - Removed " + mvObj.Name + " (Newer Move)\n";
                                    p.Move[m] = -1;
                                }
                                var rules = new GSCRules();
                                
                                if (!rules.CheckMove(p.Pokemon, mvObj)) {
                                    changeString += p.Nickname + "- Removed " + mvObj.Name + " (Not Learnable in GSC)\n";
                                    p.Move[m] = -1;
                                }
                                
                            }
                        }
                        break;
                    

                    case CompatModes.nbTrueRuSa:
                        if (!p.Pokemon.ExistADV) {
                            changeString += p.Nickname + " - Removed\n";
                            _workingSettings.Team[i] = null;
                        }
                        if (_workingSettings.CurrentCompatibilityMode == CompatModes.nbTrueRBY || 
                            _workingSettings.CurrentCompatibilityMode == CompatModes.nbRBYTrade ||
                            _workingSettings.CurrentCompatibilityMode == CompatModes.nbTrueGSC || 
                            _workingSettings.CurrentCompatibilityMode == CompatModes.nbGSCTrade) {
                            p.UpgradeToAdvStats();
                        }
                        // -- Remove GSC Only items
                        if (!p.Item.IsAdvanceItem()) {
                            changeString += p.Nickname + " - GSC-Only Item removed\n";
                            p.Item = Items.nbNoItem;
                        }
                        var rulesAgainStupidScopeBullshit = new ADVRules();
                        for (var m = 0; m < 4; m++) {
                            if (p.Move[m] > 0) {
                                var mvObj = MoveDatabase.Moves[p.Move[m]];
                                if (!rulesAgainStupidScopeBullshit.CheckMove(p.Pokemon, mvObj)) {
                                    changeString += p.Nickname + " - Removed " + mvObj.Name + " (Not Learnable in RuSa)\n";
                                    p.Move[m] = -1;
                                }
                            }
                        }

                        break;
                }

                if (p != null) p.GameVersion = newCompatMode;
            }

            if (!string.IsNullOrEmpty(changeString)) {
                MessageBox.Show(changeString, "Team Changes", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
            _workingSettings.CurrentCompatibilityMode = newCompatMode;
        }

        private void txtNickname_TextChanged(object sender, EventArgs e) {
            if (_switching)
                return;
            
            if (_currentPokemon != null && txtNickname.Text != null && txtNickname.Text != _currentPokemon.Nickname)
                _currentPokemon.Nickname = txtNickname.Text;
        }

        private void dropItem_SelectedIndexChanged(object sender, EventArgs e) {
            if (_currentPokemon != null) {
                _currentPokemon.Item = (Items)Enum.Parse(typeof(Items), "nb" + dropItem.Text.Replace(" ", ""));
                RefreshCompatTree();
            }
        }

        private void fullAdvancedToolStripMenuItem_Click(object sender, EventArgs e) {
            DoVersionChange(CompatModes.nbFullAdvance);
            LoadPokemonSlot(_currentSlot);
            UpdateSelections();
            RefreshCompatTree();
        }

        private void ruSaOnlyToolStripMenuItem_Click(object sender, EventArgs e) {
            DoVersionChange(CompatModes.nbTrueRuSa);
            LoadPokemonSlot(_currentSlot);
            UpdateSelections();
            RefreshCompatTree();
        }

        private void gSCWithTradesToolStripMenuItem_Click(object sender, EventArgs e) {
            DoVersionChange(CompatModes.nbGSCTrade);
            LoadPokemonSlot(_currentSlot);
            UpdateSelections();
            RefreshCompatTree();
        }

        private void trueGSCToolStripMenuItem_Click(object sender, EventArgs e) {
            DoVersionChange(CompatModes.nbTrueGSC);
            LoadPokemonSlot(_currentSlot);
            UpdateSelections();
            RefreshCompatTree();
        }

        private void rBYWithTradesToolStripMenuItem_Click(object sender, EventArgs e) {
            DoVersionChange(CompatModes.nbRBYTrade);
            LoadPokemonSlot(_currentSlot);
            UpdateSelections();
            RefreshCompatTree();
        }

        private void trueRBYToolStripMenuItem_Click(object sender, EventArgs e) {
            DoVersionChange(CompatModes.nbTrueRBY);
            LoadPokemonSlot(_currentSlot);
            UpdateSelections();
            RefreshCompatTree();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e) {
            ofdPnb.ShowDialog();
        }

        private void ofdPnb_FileOk(object sender, System.ComponentModel.CancelEventArgs e) {
            if (!File.Exists(ofdPnb.FileName))
                return;

            if (ofdPnb.FileName.EndsWith(".pnb")) {
                PnbFile save = new PnbFile(ofdPnb.FileName);
                save.Load();
                _workingSettings.CurrentCompatibilityMode = save.Team.FirstOrDefault(a => a != null)?.GameVersion ?? _workingSettings.CurrentCompatibilityMode;
                _workingSettings.Team = save.Team;
                _workingSettings.IconUsed = save.CurrentPicture;
                _workingSettings.CurrentGraphicsMode = (GraphicsMode)(save.GraphicsVersion);
                txtUsername.Text = save.Name;
                txtWinMessage.Text = save.WinMessage;
                txtLoseMessage.Text = save.LoseMessage;
                txtExtraInfo.Text = save.ExtraInfo;
                listView1.Items[save.CurrentPicture - 1].Selected = true;
                _workingSettings.Username = save.Name;
                _workingSettings.WinMessage = save.WinMessage;
                _workingSettings.LoseMessage = save.LoseMessage;
                _workingSettings.MoreInfo = save.ExtraInfo;

                LoadPokemonSlot(_currentSlot);
                UpdateSelections();
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e) {
            if (listView1.SelectedIndices.Count == 0)
                return;

            if (_workingSettings == null)
                return;
            
            _workingSettings.IconUsed = (byte)(listView1.SelectedIndices[0] + 1);
        }

        private void btnSwitch_Click(object sender, EventArgs e) {
            _currentPokemon.GameVersion = _workingSettings.CurrentCompatibilityMode;
            
            var expertForm = new AdvExpert(this.MdiParent, _currentPokemon);
            this.Enabled = false;
            expertForm.FormClosing += ExpertFormOnFormClosing;
            expertForm.FormClosed += ExpertFormOnFormClosed;
            expertForm.Show();
        }

        private void ExpertFormOnFormClosed(object sender, FormClosedEventArgs e) {
            AdvExpert expertForm = (AdvExpert)sender;

            if (expertForm.OkClicked) {
                _currentPokemon = expertForm.workingPoke;
                UpdateStatLabels();
                UpdateTeamEvs();
                LoadPokemonGraphic();
                PopulateMoves();
                
            }

            this.Enabled = true;
        }

        private void ExpertFormOnFormClosing(object sender, FormClosingEventArgs e) {
            this.Enabled = true;
        }

    }
}