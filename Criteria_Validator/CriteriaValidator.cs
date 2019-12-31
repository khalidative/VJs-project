using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Criteria_Validator
{
    public partial class CriteriaValidator : Form
    {
        //---------------Inputs
        string unit;
        string material;
        double TQ;
        double CS;
        double LT;
        double ALT;
        double ADC;
        double MLT;
        double MDC;
        double ROP;
        double DC;
        //---------------Outputs
        double RSDEF;
        double SSDEF;

        public CriteriaValidator()
        {
            InitializeComponent();
        }

        private void btn_Validate_Click(object sender, EventArgs e)
        {
            unit = cmb_Input_Units.SelectedItem.ToString();
            material = cmb_Input_Material.SelectedItem.ToString();
            TQ = Convert.ToDouble(txt_Input_TQ.Text);
            CS = Convert.ToDouble(txt_Input_CS.Text);
            LT = Convert.ToDouble(txt_Input_LT.Text);
            ALT = Convert.ToDouble(txt_Input_ALT.Text);
            ADC = Convert.ToDouble(txt_Input_ADC.Text);
            MLT = Convert.ToDouble(txt_Input_MLT.Text);
            MDC = Convert.ToDouble(txt_Input_MDC.Text);
            ROP = Convert.ToDouble(txt_Input_ROP.Text);
            DC = Convert.ToDouble(txt_Input_DC.Text);

            //RSDEF
            if(DC > ADC)
            {
                RSDEF = (DC - ADC) * ALT;
            }
            else
            {
                RSDEF = 0;
            }

            //SSDEF
            if(LT > ALT)
            {
                SSDEF = (LT - ALT) * ADC;
            }
            else
            {
                SSDEF = 0;
            }


            lbl_Output_RSDEF.Text = RSDEF.ToString();
            lbl_Output_SSDEF.Text = SSDEF.ToString();

        }

    }
}
