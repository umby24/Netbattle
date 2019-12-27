using Netbattle.Common;
using System;
using System.Windows.Forms;
using Netbattle.Database;

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
            PokemonDatabase.Load();
            Configuration.Load();
            UserSettings.Load();
        }
        

        private void mehToolStripMenuItem_Click(object sender, EventArgs e) {
            _qs.Show();
        }
    }
}
