using System.ComponentModel;
using Netbattle.Core.Battle;
using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace Netbattle.Forms {
    public partial class AdvExpert : Form
    {
        public TeamMember workingPoke;
        public bool OkClicked = false;

        public AdvExpert(Form parent, TeamMember pokemon)
        {
            InitializeComponent();
            MdiParent = parent;
            workingPoke = pokemon;
            for(int i = 2; i < 17; i++) { 
                cmbHiddenPower.Items.Add(Elements.GetName(typeof(Elements), i));
            }
            LoadPokemon();
        }

        public void LoadPokemon()
        {
            this.Text = "Expert Mode: " + workingPoke.Pokemon.Name;

            chkShiny.Checked = workingPoke.Shiny;
            trkAtk.Value = workingPoke.EVs[StatTypes.nbAttack];
            trkDef.Value = workingPoke.EVs[StatTypes.nbDefense];
            trkSpd.Value = workingPoke.EVs[StatTypes.nbSpeed];
            trkSpDef.Value = workingPoke.EVs[StatTypes.nbSpDefense];
            trkHp.Value = workingPoke.EVs[StatTypes.nbHP];
            TrkSpAtk.Value = workingPoke.EVs[StatTypes.nbSpAttack];

            cmbIVAtk.SelectedIndex = Math.Abs(31 - workingPoke.DVs[StatTypes.nbAttack]);
            cmbIVDef.SelectedIndex = Math.Abs(31 - workingPoke.DVs[StatTypes.nbDefense]);
            cmbIVSpd.SelectedIndex = Math.Abs(31 - workingPoke.DVs[StatTypes.nbSpeed]);
            cmbIVSpDef.SelectedIndex = Math.Abs(31 - workingPoke.DVs[StatTypes.nbSpDefense]);
            cmbIVHP.SelectedIndex = Math.Abs(31 - workingPoke.DVs[StatTypes.nbHP]);
            cmbIVSpAtk.SelectedIndex = Math.Abs(31 - workingPoke.DVs[StatTypes.nbSpAttack]);

            cmbNature.SelectedIndex = (workingPoke.NatureNum == 255) ? 0 : workingPoke.NatureNum + 1;
            numericUpDown1.Value = workingPoke.Level;

            RefreshStats();
            RefreshEvLabels();
            RefreshEVBar();
            if (workingPoke.Pokemon.No == 201)
            {
                cmbUnknownLetter.Visible = true;
                lblUnkLetter.Visible = true;
                cmbUnknownLetter.SelectedIndex = workingPoke.UnownLetter;
            }
            else
            {
                cmbUnknownLetter.Visible = false;
                lblUnkLetter.Visible = false;
            }

            if (workingPoke.GameVersion == CompatModes.nbModAdv)
            {
                radioButton1.Text = TypeDatabase.AttributeText[(int)workingPoke.Pokemon.ModAttr[0]];
                radioButton2.Text = TypeDatabase.AttributeText[(int)workingPoke.Pokemon.ModAttr[1]];
            }
            else
            {
                radioButton1.Text = TypeDatabase.AttributeText[(int)workingPoke.Pokemon.PAtt[0]];
                radioButton2.Text = TypeDatabase.AttributeText[(int)workingPoke.Pokemon.PAtt[1]];
            }

            if (workingPoke.Pokemon.PercentFemale == -1)
            {
                // -- Genderless :) 
                lblGender.Enabled = false;
                radioMale.Enabled = false;
                radioFemale.Enabled = false;
            }
            else
            {
                lblGender.Enabled = true;
                radioMale.Enabled = true;
                radioFemale.Enabled = true;
                if (workingPoke.Gender == (byte)Gender.Male)
                {
                    radioMale.Checked = true;
                    radioFemale.Checked = false;
                }
                else if (workingPoke.Gender == (byte)Gender.Female)
                {
                    radioFemale.Checked = true;
                    radioMale.Checked = false;
                }
            }
        }

        private int EVTotal()
        {
            int result = 0;
            result += trkAtk.Value;
            result += trkDef.Value;
            result += trkSpd.Value;
            result += trkSpDef.Value;
            result += trkHp.Value;
            result += TrkSpAtk.Value;
            return result;
        }

        private void RefreshEVBar()
        {
            int currentTotal = EVTotal();
            lblRemaining.Text = $"Remaining EP: {510 - currentTotal}";
            float percent = ((510 - currentTotal) / 510f) * 100f;
            prgTotal.Value = (int)percent;
        }

        private void RefreshEvLabels()
        {
            lblHpVal.Text = trkHp.Value.ToString();
            lblAtkVal.Text = trkAtk.Value.ToString();
            lblDefVal.Text = trkDef.Value.ToString();
            lblSpdVal.Text = trkSpd.Value.ToString();
            lblSpAtkVal.Text = TrkSpAtk.Value.ToString();
            lblSpDefVal.Text = trkSpDef.Value.ToString();
        }

        private void RefreshStats()
        {
            workingPoke.EVs = new StatSet()
            {
                [StatTypes.nbHP] = trkHp.Value,
                [StatTypes.nbAttack] = trkAtk.Value,
                [StatTypes.nbDefense] = trkDef.Value,
                [StatTypes.nbSpeed] = trkSpd.Value,
                [StatTypes.nbSpAttack] = TrkSpAtk.Value,
                [StatTypes.nbSpDefense] = trkSpDef.Value
            };
            workingPoke.EVs = new StatSet()
            {
                [StatTypes.nbHP] = trkHp.Value,
                [StatTypes.nbAttack] = trkAtk.Value,
                [StatTypes.nbDefense] = trkDef.Value,
                [StatTypes.nbSpeed] = trkSpd.Value,
                [StatTypes.nbSpAttack] = TrkSpAtk.Value,
                [StatTypes.nbSpDefense] = trkSpDef.Value
            };

            workingPoke.NatureNum = (byte)cmbNature.SelectedIndex;
            workingPoke.Level = (byte)numericUpDown1.Value;

            workingPoke.Pokemon.MaxHP = BattleSystem.GetGenerationHp(CompatModes.nbFullAdvance,
                workingPoke.Pokemon.BaseStats[StatTypes.nbHP], workingPoke.DVs[StatTypes.nbHP],
                workingPoke.EVs[StatTypes.nbHP], workingPoke.Level);

            workingPoke.CurrentStats[StatTypes.nbHP] = workingPoke.Pokemon.MaxHP;

            workingPoke.CurrentStats[StatTypes.nbAttack] = BattleSystem.GetGenerationStat(CompatModes.nbFullAdvance,
                workingPoke.Pokemon.BaseStats[StatTypes.nbAttack],
                workingPoke.DVs[StatTypes.nbAttack], workingPoke.EVs[StatTypes.nbAttack],
                workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 1));
            workingPoke.CurrentStats[StatTypes.nbDefense] = BattleSystem.GetGenerationStat(CompatModes.nbFullAdvance,
                workingPoke.Pokemon.BaseStats[StatTypes.nbDefense],
                workingPoke.DVs[StatTypes.nbDefense], workingPoke.EVs[StatTypes.nbDefense],
                workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 2));
            workingPoke.CurrentStats[StatTypes.nbSpAttack] = BattleSystem.GetGenerationStat(CompatModes.nbFullAdvance,
                workingPoke.Pokemon.BaseStats[StatTypes.nbSpAttack],
                workingPoke.DVs[StatTypes.nbSpAttack], workingPoke.EVs[StatTypes.nbSpAttack],
                workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 3));
            workingPoke.CurrentStats[StatTypes.nbSpDefense] = BattleSystem.GetGenerationStat(CompatModes.nbFullAdvance,
                workingPoke.Pokemon.BaseStats[StatTypes.nbSpDefense],
                workingPoke.DVs[StatTypes.nbSpDefense], workingPoke.EVs[StatTypes.nbSpDefense],
                workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 4));
            workingPoke.CurrentStats[StatTypes.nbSpeed] = BattleSystem.GetGenerationStat(CompatModes.nbFullAdvance,
                workingPoke.Pokemon.BaseStats[StatTypes.nbSpeed],
                workingPoke.DVs[StatTypes.nbSpeed], workingPoke.EVs[StatTypes.nbSpeed],
                workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 5));

            lblFinalHp.Text = workingPoke.CurrentStats[StatTypes.nbHP].ToString();
            lblFinalAtk.Text = workingPoke.CurrentStats[StatTypes.nbAttack].ToString();
            lblFinalDef.Text = workingPoke.CurrentStats[StatTypes.nbDefense].ToString();
            lblFinalSpAtk.Text = workingPoke.CurrentStats[StatTypes.nbSpAttack].ToString();
            lblFinalSpDef.Text = workingPoke.CurrentStats[StatTypes.nbSpDefense].ToString();
            lblFinalSpd.Text = workingPoke.CurrentStats[StatTypes.nbSpeed].ToString();
            IGenerationRuleset rules = BattleSystem.GetGenerationRuleset(workingPoke.GameVersion);
            // lbl HiddenPower
            lblHiddenPower.Text = "Hidden Power: " + rules.HiddenPowerStrength(workingPoke.DVs).ToString();
            var type = rules.HiddenPowerType(workingPoke.DVs);
            cmbHiddenPower.SelectedIndex = (int)type - 2;
        }

        private void chkShiny_CheckedChanged(object sender, EventArgs e)
        {
            workingPoke.Shiny = chkShiny.Checked;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OkClicked = true;
            this.Close();
        }

        private void btnEVClear_Click(object sender, EventArgs e)
        {
            trkAtk.Value = 0;
            trkDef.Value = 0;
            trkSpd.Value = 0;
            trkSpDef.Value = 0;
            trkHp.Value = 0;
            TrkSpAtk.Value = 0;
            RefreshEVBar();
            RefreshEvLabels();
            RefreshStats();
        }

        private void trkHp_Scroll(object sender, EventArgs e)
        {
            if (chkLkHp.Checked)
            {
                trkHp.Value = workingPoke.EVs[StatTypes.nbHP];
                return;
            }

            if (chkSnap.Checked)
                trkHp.Value = (trkHp.Value / 4) * 4;

            if ((510 - EVTotal()) < 0)
            {
                trkHp.Value = workingPoke.EVs[StatTypes.nbHP];
                return;
            }

            workingPoke.EVs[StatTypes.nbHP] = (byte)trkHp.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkAtk_Scroll(object sender, EventArgs e)
        {
            if (chkLkAtk.Checked)
            {
                trkAtk.Value = workingPoke.EVs[StatTypes.nbAttack];
                return;
            }

            if (chkSnap.Checked)
                trkAtk.Value = (trkAtk.Value / 4) * 4;

            if ((510 - EVTotal()) < 0)
            {
                trkAtk.Value = workingPoke.EVs[StatTypes.nbAttack];
                return;
            }

            workingPoke.EVs[StatTypes.nbAttack] = (byte)trkAtk.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkDef_Scroll(object sender, EventArgs e)
        {
            if (chkLkDef.Checked)
            {
                trkDef.Value = workingPoke.EVs[StatTypes.nbDefense];
                return;
            }

            if (chkSnap.Checked)
                trkDef.Value = (trkDef.Value / 4) * 4;

            if ((510 - EVTotal()) < 0)
            {
                trkDef.Value = workingPoke.EVs[StatTypes.nbDefense];
                return;
            }

            workingPoke.EVs[StatTypes.nbDefense] = (byte)trkDef.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkSpd_Scroll(object sender, EventArgs e)
        {
            if (chkLkSpd.Checked)
            {
                trkSpd.Value = workingPoke.EVs[StatTypes.nbSpeed];
                return;
            }

            if (chkSnap.Checked)
                trkSpd.Value = (trkSpd.Value / 4) * 4;

            if ((510 - EVTotal()) < 0)
            {
                trkSpd.Value = workingPoke.EVs[StatTypes.nbSpeed];
                return;
            }

            workingPoke.EVs[StatTypes.nbSpeed] = (byte)trkSpd.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void TrkSpAtk_Scroll(object sender, EventArgs e)
        {
            if (chkLkSpAtk.Checked)
            {
                TrkSpAtk.Value = workingPoke.EVs[StatTypes.nbSpAttack];
                return;
            }

            if (chkSnap.Checked)
                TrkSpAtk.Value = (TrkSpAtk.Value / 4) * 4;

            if ((510 - EVTotal()) < 0)
            {
                TrkSpAtk.Value = workingPoke.EVs[StatTypes.nbSpAttack];
                return;
            }

            workingPoke.EVs[StatTypes.nbSpAttack] = (byte)TrkSpAtk.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkSpDef_Scroll(object sender, EventArgs e)
        {
            if (chkLkSpDf.Checked)
            {
                trkSpDef.Value = workingPoke.EVs[StatTypes.nbSpDefense];
                return;
            }

            if (chkSnap.Checked)
                trkSpDef.Value = (trkSpDef.Value / 4) * 4;

            if ((510 - EVTotal()) < 0)
            {
                trkSpDef.Value = workingPoke.EVs[StatTypes.nbSpDefense];
                return;
            }

            workingPoke.EVs[StatTypes.nbSpDefense] = (byte)trkSpDef.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void chkSnap_CheckedChanged(object sender, EventArgs e)
        {
            int y = (chkSnap.Checked ? 4 : 1);
            trkHp.SmallChange = y;
            trkHp.SmallChange = y;
            trkDef.SmallChange = y;
            trkSpd.SmallChange = y;
            trkSpDef.SmallChange = y;
            TrkSpAtk.SmallChange = y;
        }


        private void cmbHiddenPower_SelectedIndexChanged(object sender, EventArgs e)
        {
            int[] modifiers = new int[6];

            switch (cmbHiddenPower.SelectedIndex)
            {
                case 0:
                    modifiers[1] = 1;
                    modifiers[2] = 1;
                    modifiers[3] = 1;
                    break;
                case 1:
                    modifiers[2] = 1;
                    modifiers[3] = 1;
                    break;
                case 2:
                    modifiers[3] = 1;
                    break;
                case 3:
                    modifiers[1] = 1;
                    modifiers[3] = 1;
                    break;
                case 4:
                    modifiers[2] = 1;
                    break;
                case 5:
                    modifiers[1] = 1;
                    modifiers[2] = 1;
                    modifiers[3] = 1;
                    modifiers[4] = 1;
                    break;
                case 6:
                    modifiers[1] = 1;
                    modifiers[3] = 1;
                    modifiers[4] = 1;
                    break;
                case 7:
                    modifiers[3] = 1;
                    modifiers[4] = 1;
                    break;
                case 8:
                    modifiers[2] = 1;
                    modifiers[3] = 1;
                    modifiers[4] = 1;
                    break;
                case 9:
                    modifiers[1] = 1;
                    modifiers[2] = 1;
                    break;
                case 10:
                    modifiers[2] = 1;
                    modifiers[4] = 1;
                    break;
                case 11:
                    modifiers[1] = 1;
                    modifiers[2] = 1;
                    modifiers[4] = 1;
                    break;
                case 12:
                    modifiers[1] = 1;
                    modifiers[4] = 1;
                    break;
                case 13:
                    modifiers[1] = 1;
                    break;
                case 14:
                case 15:
                    modifiers[4] = 1;
                    break;
            }

            cmbIVHP.SelectedIndex = modifiers[0];
            cmbIVAtk.SelectedIndex = modifiers[1];
            cmbIVDef.SelectedIndex = modifiers[2];
            cmbIVSpd.SelectedIndex = modifiers[3];
            cmbIVSpAtk.SelectedIndex = modifiers[4];
            cmbIVSpDef.SelectedIndex = modifiers[5];

        }

        private void cmbNature_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbNature.SelectedIndex % 5 == cmbNature.SelectedIndex / 5)
            {
                lblMinus.Visible = false;
                lblPlus.Visible = false;
                return;
            }
            // -- 94, 26: First label location.

            lblPlus.Visible = true;
            lblMinus.Visible = true;
            for (int i = 1; i < 6; i++)
            {
                if (BattleSystem.NatureStats[cmbNature.SelectedIndex].StatChg[i] == 1)
                {
                    lblPlus.Location = new Point(121, 26 + (i * 26));
                }
                else if (BattleSystem.NatureStats[cmbNature.SelectedIndex].StatChg[i] == -1)
                {
                    lblMinus.Location = new Point(121, 26 + (i * 26));
                }
            }

            RefreshStats();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmbIVHP_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshStats();
        }

        private void cmbIVAtk_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshStats();
        }

        private void cmbIVDef_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshStats();
        }

        private void cmbIVSpd_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshStats();
        }

        private void cmbIVSpAtk_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshStats();
        }

        private void cmbIVSpDef_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshStats();
        }

        private void cmbUnknownLetter_SelectedIndexChanged(object sender, EventArgs e)
        {
            workingPoke.UnownLetter = (byte)cmbUnknownLetter.SelectedIndex;
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            chkLkAtk.Checked = true;
            chkLkDef.Checked = true;
            chkLkSpAtk.Checked = true;
            chkLkSpDf.Checked = true;
            chkLkHp.Checked = true;
            chkLkSpd.Checked = true;
        }

        private void btnUnlock_Click(object sender, EventArgs e)
        {
            chkLkAtk.Checked = false;
            chkLkDef.Checked = false;
            chkLkSpAtk.Checked = false;
            chkLkSpDf.Checked = false;
            chkLkHp.Checked = false;
            chkLkSpd.Checked = false;
        }

        private void radioMale_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioFemale_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            workingPoke.Level = (byte)numericUpDown1.Value;
            RefreshStats();
        }
    }
}