using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Network;
using Netbattle.Network;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Netbattle.Forms {
    public partial class ChallengeWindow : Form
    {
        // -- State Flags.
        public bool ICalled;
        public bool PressedOk;
        public int GameMode;
        private NbClient _client;
        private Player _challenger;
        private TerrainTypes _tTypes;
        private BattleModes _cMode;
        private byte[] rules;

        /// <summary>
        /// Used when viewing / challenging someone else.
        /// </summary>
        /// <param name="nbClient"></param>
        /// <param name="challenger"></param>
        /// <param name="parent"></param>
        public ChallengeWindow(NbClient nbClient, Player challenger, Form parent)
        {
            this.MdiParent = parent.MdiParent;
            InitializeComponent();
            _client = nbClient;
            _challenger = challenger;
            this.Top = (Screen.PrimaryScreen.Bounds.Height - this.Height) / 2;
            ICalled = true;
        }

        /// <summary>
        /// Used when someone actually challenges us
        /// </summary>
        /// <param name="nbClient"></param>
        /// <param name="challenger"></param>
        /// <param name="parent"></param>
        public ChallengeWindow(NbClient nbClient, Player challenger, Form parent, TerrainTypes terrain, BattleModes challengeMode, byte[] ruleSet)
        {
            this.MdiParent = parent.MdiParent;
            InitializeComponent();
            _client = nbClient;
            _challenger = challenger;
            _tTypes = terrain;
            _cMode = challengeMode;
            rules = ruleSet;
            this.Top = (Screen.PrimaryScreen.Bounds.Height - this.Height) / 2;
        }

        private void PopulateForm()
        {
            if (ICalled)
            {
                label4.Text = "Challenge Setup";
            }
            else
            {
                btnChallenge.Text = "Accept";
            }
            progressBar1.Value = _challenger.TeamPower;
            label1.Text = _challenger.Description;
            if (_challenger.ShowTeam)
            {
                var pictureBoxes = new[] { pk1, pk2, pk3, pk4, pk5, pk6 };
                var i = 0;
                foreach (var poke in _challenger.Team)
                {
                    byte[] img = GraphicsDatabase.GetSprite(poke, GraphicsMode.nbGFXSml, false);
                    using (var ms = new MemoryStream(img))
                    {
                        pictureBoxes[i].Image = Image.FromStream(ms);
                    }

                    i++;
                }
            }
            else
            {
                var img = GraphicsDatabase.GetSprite("000rs");
                Image actual;
                using (var ms = new MemoryStream(img))
                {
                    actual = Image.FromStream(ms);
                }
                var pictureBoxes = new[] { pk1, pk2, pk3, pk4, pk5, pk6 };

                for (int i = 0; i < 6; i++)
                {

                    pictureBoxes[i].Image = actual;
                    pictureBoxes[i].SizeMode = PictureBoxSizeMode.StretchImage;
                }
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (ICalled)
            {
                ICalled = false;
                if (PressedOk)
                {
                    var cancelPacket = new ChallengedCancelled()
                    {
                        ChallengeNumber = (byte)_challenger.Id
                    };
                    _client.SendPacket(cancelPacket);
                    _client.Battling = false;
                }

                Close();
                return;
            }

            var refusedPacket = new PlayerRefusedBattle()
            {
                ChallengeNumber = (byte)_challenger.Id
            };
            _client.SendPacket(refusedPacket);
            _client.Battling = false;
            Close();
        }

        private void ChallengeWindow_Load(object sender, EventArgs e)
        {
            PopulateForm();
        }
    }
}
