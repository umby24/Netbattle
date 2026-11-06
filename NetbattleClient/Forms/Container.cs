using Netbattle;
using Netbattle.Common;
using Netbattle.Core.Database;

namespace NetbattleClient.Forms {
    public partial class Container : Form {
        private readonly Form1 _qs;

        public Container() {
            InitializeComponent();
            _qs = new Form1(this);
        }

        private void Container_Load(object sender, EventArgs e) {
            MoveDatabase.Load();
            TypeDatabase.Load();
            GraphicsDatabase.Load();
            GraphicsDatabase.LoadNbs();
            PokemonDatabase.Load();
            BattleSystem.GenerateStatTable();
            Configuration.Load();
            UserSettings.Load(Configuration.CurrentSettings.LastJnb, Configuration.CurrentSettings.LastPnb);
            BattleSystem.GenerateStatTable();
            _qs.Show();
        }
    }
}
