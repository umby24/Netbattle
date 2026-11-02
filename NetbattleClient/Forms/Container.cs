using Netbattle.Common;
using System;
using System.Windows.Forms;
using Netbattle.Core.Database;
using NetbattleClient.Forms;

namespace Netbattle.Forms {
    public partial class Container : Form {
        private Form1 _qs;

        public Container() {
            InitializeComponent();
        }

        private void Container_Load(object sender, EventArgs e) {
            _qs = new Form1(this);
            _qs.Show();


            MoveDatabase.Load();
            TypeDatabase.Load();
            GraphicsDatabase.Load();
            GraphicsDatabase.LoadNbs();
            PokemonDatabase.Load();
            BattleSystem.GenerateStatTable();
            Configuration.Load();
            UserSettings.Load(Configuration.CurrentSettings.LastJnb, Configuration.CurrentSettings.LastPnb);
            BattleSystem.GenerateStatTable();
        }
        

        private void mehToolStripMenuItem_Click(object sender, EventArgs e) {
            _qs.Show();
        }
    }
}
