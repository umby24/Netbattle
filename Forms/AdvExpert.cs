using System;
using System.Drawing;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Database;

namespace Netbattle.Forms {
    public partial class AdvExpert : Form {
        public Pokemon workingPoke;
        public bool OkClicked = false;

        public AdvExpert(Pokemon pokemon) {
            InitializeComponent();
            workingPoke = pokemon;
            LoadPokemon();
        }

        public void LoadPokemon() {
            this.Text = "Expert Mode: " + workingPoke.Name;

            chkShiny.Checked = workingPoke.Shiny;
            trkAtk.Value = workingPoke.EV_Atk;
            trkDef.Value = workingPoke.EV_Def;
            trkSpd.Value = workingPoke.EV_Spd;
            trkSpDef.Value = workingPoke.EV_SDef;
            trkHp.Value = workingPoke.EV_HP;
            TrkSpAtk.Value = workingPoke.EV_SAtk;

            cmbIVAtk.SelectedIndex = Math.Abs(31-workingPoke.DV_Atk);
            cmbIVDef.SelectedIndex = Math.Abs(31-workingPoke.DV_Def);
            cmbIVSpd.SelectedIndex = Math.Abs(31-workingPoke.DV_Spd);
            cmbIVSpDef.SelectedIndex = Math.Abs(31-workingPoke.DV_SDef);
            cmbIVHP.SelectedIndex = Math.Abs(31-workingPoke.DV_HP);
            cmbIVSpAtk.SelectedIndex = Math.Abs(31-workingPoke.DV_SAtk);

            cmbNature.SelectedIndex = workingPoke.NatureNum;
            numericUpDown1.Value = workingPoke.Level;

            RefreshStats();
            RefreshEvLabels();
            RefreshEVBar();
            if (workingPoke.No == 201) {
                cmbUnknownLetter.Visible = true;
                lblUnkLetter.Visible = true;
                cmbUnknownLetter.SelectedIndex = workingPoke.UnownLetter;
            }
            else {
                cmbUnknownLetter.Visible = false;
                lblUnkLetter.Visible = false;
            }

            if (workingPoke.GameVersion == CompatModes.nbModAdv) {
                radioButton1.Text = TypeDatabase.AttributeText[(int)workingPoke.ModAttr[0]];
                radioButton2.Text = TypeDatabase.AttributeText[(int)workingPoke.ModAttr[1]];
            }
            else {
                radioButton1.Text = TypeDatabase.AttributeText[(int)workingPoke.PAtt[0]];
                radioButton2.Text = TypeDatabase.AttributeText[(int)workingPoke.PAtt[1]];
            }

            if (workingPoke.PercentFemale == -1) {
                // -- Genderless :) 
                lblGender.Enabled = false;
                radioMale.Enabled = false;
                radioFemale.Enabled = false;
            }
            else {
                lblGender.Enabled = true;
                radioMale.Enabled = true;
                radioFemale.Enabled = true;
                if (workingPoke.Gender == (byte)Gender.Male) {
                    radioMale.Checked = true;
                    radioFemale.Checked = false;
                }
                else if (workingPoke.Gender == (byte)Gender.Female) {
                    radioFemale.Checked = true;
                    radioMale.Checked = false;
                }
            }
        }

        private int EVTotal() {
            int result = 0;
            result += trkAtk.Value;
            result += trkDef.Value;
            result += trkSpd.Value;
            result += trkSpDef.Value;
            result += trkHp.Value;
            result += TrkSpAtk.Value;
            return result;
        }

        private void RefreshEVBar() {
            int currentTotal = EVTotal();
            lblRemaining.Text = $"Remaining EP: {510 - currentTotal}";
            float percent = ((510 - currentTotal) / 510f) * 100f;
            prgTotal.Value = (int)percent;
        }

        private void RefreshEvLabels() {
            lblHpVal.Text = trkHp.Value.ToString();
            lblAtkVal.Text = trkAtk.Value.ToString();
            lblDefVal.Text = trkDef.Value.ToString();
            lblSpdVal.Text = trkSpd.Value.ToString();
            lblSpAtkVal.Text = TrkSpAtk.Value.ToString();
            lblSpDefVal.Text = trkSpDef.Value.ToString();
        }

        private void RefreshStats() {
            workingPoke.DV_HP = (byte)cmbIVHP.SelectedIndex;
            workingPoke.DV_Atk = (byte)cmbIVAtk.SelectedIndex;
            workingPoke.DV_Def = (byte)cmbIVDef.SelectedIndex;
            workingPoke.DV_SAtk = (byte)cmbIVSpAtk.SelectedIndex;
            workingPoke.DV_SDef = (byte)cmbIVSpDef.SelectedIndex;
            workingPoke.DV_Spd = (byte)cmbIVSpd.SelectedIndex;
            workingPoke.EV_Atk = (byte)trkAtk.Value;
            workingPoke.EV_Def = (byte)trkDef.Value;
            workingPoke.EV_Spd = (byte)trkSpd.Value;
            workingPoke.EV_SDef = (byte)trkSpDef.Value;
            workingPoke.EV_HP = (byte)trkHp.Value;
            workingPoke.EV_SAtk = (byte)TrkSpAtk.Value;
            workingPoke.NatureNum = (byte)cmbNature.SelectedIndex;
            workingPoke.Level = (byte)numericUpDown1.Value;
            workingPoke.MaxHP = BattleSystem.GetAdvHp(workingPoke.BaseHP, workingPoke.DV_HP, workingPoke.EV_HP,
                workingPoke.Level);
            workingPoke.HP = workingPoke.MaxHP;
            workingPoke.Attack = BattleSystem.GetAdvStat(workingPoke.BaseAttack, workingPoke.DV_Atk, workingPoke.EV_Atk,
                workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 1));
            workingPoke.Defense = BattleSystem.GetAdvStat(workingPoke.BaseDefense, workingPoke.DV_Def,
                workingPoke.EV_Def, workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 2));
            workingPoke.SpecialAttack = BattleSystem.GetAdvStat(workingPoke.BaseSAttack, workingPoke.DV_SAtk,
                workingPoke.EV_SAtk, workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 3));
            workingPoke.SpecialDefense = BattleSystem.GetAdvStat(workingPoke.BaseSDefense, workingPoke.DV_SDef,
                workingPoke.EV_SDef, workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 4));
            workingPoke.Speed = BattleSystem.GetAdvStat(workingPoke.BaseSpeed, workingPoke.DV_Spd, workingPoke.EV_Spd,
                workingPoke.Level, NbMethods.GetNatureMod(workingPoke.NatureNum, 5));

            lblFinalHp.Text = workingPoke.HP.ToString();
            lblFinalAtk.Text = workingPoke.Attack.ToString();
            lblFinalDef.Text = workingPoke.Defense.ToString();
            lblFinalSpAtk.Text = workingPoke.SpecialAttack.ToString();
            lblFinalSpDef.Text = workingPoke.SpecialDefense.ToString();
            lblFinalSpd.Text = workingPoke.Speed.ToString();
            // lbl HiddenPower
        }

        private void chkShiny_CheckedChanged(object sender, EventArgs e) {
            workingPoke.Shiny = chkShiny.Checked;
        }

        private void button4_Click(object sender, EventArgs e) {
            OkClicked = true;
            this.Close();
        }

        private void btnEVClear_Click(object sender, EventArgs e) {
            trkAtk.Value = 0;
            trkDef.Value = 0;
            trkSpd.Value = 0;
            trkSpDef.Value = 0;
            trkHp.Value = 0;
            TrkSpAtk.Value = 0;
            RefreshEVBar();
        }

        private void trkHp_Scroll(object sender, EventArgs e) {
            if (chkLkHp.Checked) {
                trkHp.Value = workingPoke.EV_HP;
                return;
            }

            if (chkSnap.Checked)
                trkHp.Value = (trkHp.Value / 4) * 4;

            if ((510 - EVTotal()) < 0) {
                trkHp.Value = workingPoke.EV_HP;
                return;
            }

            workingPoke.EV_HP = (byte)trkHp.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkAtk_Scroll(object sender, EventArgs e) {
            if (chkLkAtk.Checked) {
                trkAtk.Value = workingPoke.EV_Atk;
                return;
            }

            if (chkSnap.Checked)
                trkAtk.Value = (trkAtk.Value / 4) * 4;

            if ((510 - EVTotal()) < 0) {
                trkAtk.Value = workingPoke.EV_Atk;
                return;
            }

            workingPoke.EV_Atk = (byte)trkAtk.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkDef_Scroll(object sender, EventArgs e) {
            if (chkLkDef.Checked) {
                trkDef.Value = workingPoke.EV_Def;
                return;
            }

            if (chkSnap.Checked)
                trkDef.Value = (trkDef.Value / 4) * 4;

            if ((510 - EVTotal()) < 0) {
                trkDef.Value = workingPoke.EV_Def;
                return;
            }

            workingPoke.EV_Def = (byte)trkDef.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkSpd_Scroll(object sender, EventArgs e) {
            if (chkLkSpd.Checked) {
                trkSpd.Value = workingPoke.EV_Spd;
                return;
            }

            if (chkSnap.Checked)
                trkSpd.Value = (trkSpd.Value / 4) * 4;

            if ((510 - EVTotal()) < 0) {
                trkSpd.Value = workingPoke.EV_Spd;
                return;
            }

            workingPoke.EV_Spd = (byte)trkSpd.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void TrkSpAtk_Scroll(object sender, EventArgs e) {
            if (chkLkSpAtk.Checked) {
                TrkSpAtk.Value = workingPoke.EV_SAtk;
                return;
            }

            if (chkSnap.Checked)
                TrkSpAtk.Value = (TrkSpAtk.Value / 4) * 4;

            if ((510 - EVTotal()) < 0) {
                TrkSpAtk.Value = workingPoke.EV_SAtk;
                return;
            }

            workingPoke.EV_SAtk = (byte)TrkSpAtk.Value;
            RefreshEVBar();
            RefreshStats();
            RefreshEvLabels();
        }

        private void trkSpDef_Scroll(object sender, EventArgs e) {
            throw new System.NotImplementedException();
        }

        private void chkSnap_CheckedChanged(object sender, EventArgs e) {
            int y = (chkSnap.Checked ? 4 : 1);
            trkHp.SmallChange = y;
            trkHp.SmallChange = y;
            trkDef.SmallChange = y;
            trkSpd.SmallChange = y;
            trkSpDef.SmallChange = y;
            TrkSpAtk.SmallChange = y;
        }


        private void cmbHiddenPower_SelectedIndexChanged(object sender, EventArgs e) {
            int[] modifiers = new int[6];

            switch (cmbHiddenPower.SelectedIndex) {
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

        private void cmbNature_SelectedIndexChanged(object sender, EventArgs e) {
            if (cmbNature.SelectedIndex % 5 == cmbNature.SelectedIndex / 5) {
                lblMinus.Visible = false;
                lblPlus.Visible = false;
                return;
            }
            // -- 94, 26: First label location.
            
            lblPlus.Visible = true;
            lblMinus.Visible = true;
            for (int i = 1; i < 6; i++) {
                if(BattleSystem.NatureStats[cmbNature.SelectedIndex].StatChg[i] == 1) {
                    lblPlus.Location = new Point(121, 26 + (i * 26));
                }
                else if (BattleSystem.NatureStats[cmbNature.SelectedIndex].StatChg[i] == -1) {
                    lblMinus.Location = new Point(121, 26 + (i * 26));
                }
            }
            
            RefreshStats();
        }

        private void button5_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}