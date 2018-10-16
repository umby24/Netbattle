using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Database;

namespace Netbattle.Forms {
    public partial class TeamBuilder : Form {
        private Pokemon _currentPokemon;

        public TeamBuilder(Form parent) {
            this.MdiParent = parent;
            InitializeComponent();
        }

        private void TeamBuilder_Load(object sender, EventArgs e) {
            dropGraphics.SelectedIndex = 6;

            // -- Loads the pokemon!
            var pokemonNames = PokemonDatabase.BasePokemon.Select(a => a.Name).ToArray();
            dropPokemon.Items.AddRange(pokemonNames);

            // -- Load your user info..
            txtLoseMessage.Text = UserSettings.LoseMessage;
            txtWinMessage.Text = UserSettings.WinMessage;
            txtNickname.Text = UserSettings.Username;
            txtExtraInfo.Text = UserSettings.MoreInfo;
        }

        private void dropPokemon_SelectedIndexChanged(object sender, EventArgs e) {
            var pkmnObj = PokemonDatabase.BasePokemon.FirstOrDefault(a => a.Name == (string) dropPokemon.SelectedItem);
            _currentPokemon = pkmnObj;

            var img = GraphicsDatabase.GetSprite(pkmnObj, UserSettings.CurrentGraphicsMode, false);
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
    }
}
