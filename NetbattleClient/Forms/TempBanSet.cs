using System;
using System.Windows.Forms;

namespace Netbattle.Forms
{
    public partial class TempBanSet : Form
    {
        public int Value
        {
            get
            {
                return (int)duration.Value;
            }
            private set
            {
                duration.Value = value;
            }
        }

        public TempBanSet()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}