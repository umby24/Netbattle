using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Netbattle.Core.Network;
using Netbattle.Network;

namespace Netbattle.Forms {
    public partial class ChallengeWindow : Form {
        // -- State Flags.
        public bool ICalled;
        public bool PressedOk;
        public int GameMode;
        private NbClient _client;
        private int challengeSlot;
        
        public ChallengeWindow(NbClient nbClient, int challenger) {
            InitializeComponent();
            _client = nbClient;
            challengeSlot = challenger;
            this.Top = (Screen.PrimaryScreen.Bounds.Height - this.Height) / 2;
            
        }

        private void pictureBox3_Click(object sender, EventArgs e) {

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
                        ChallengeNumber = (byte)challengeSlot
                    };
                    _client.SendPacket(cancelPacket);
                    _client.Battling = false;
                }

                Close();
                return;
            }

            var refusedPacket = new PlayerRefusedBattle()
            {
                ChallengeNumber = (byte)challengeSlot
            };
            _client.SendPacket(refusedPacket);
            _client.Battling = false;
            Close();
        }
    }
}
