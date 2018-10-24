using Netbattle.Common;
using System;
using System.Windows.Forms;
using Netbattle.Database;

namespace Netbattle.Forms {
    public partial class Container : Form {
        public Container() {
            InitializeComponent();
        }

        private void Container_Load(object sender, EventArgs e) {
            var qs = new Form1(this);
            qs.FormClosed += QsOnFormClosed;
            qs.Show();


            MoveDatabase.Load();
            TypeDatabase.Load();
            GraphicsDatabase.Load();
            PokemonDatabase.Load();
            Configuration.Load();
            UserSettings.Load();
            new PnbFile("D:\\Documents\\myteam.pnb").Load();
        }

        private void QsOnFormClosed(object sender, FormClosedEventArgs formClosedEventArgs) {
            Console.WriteLine("THE QUICK START CLOSED!");
        }
    }
}
