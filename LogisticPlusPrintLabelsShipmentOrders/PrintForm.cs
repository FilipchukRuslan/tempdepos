using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LogisticPlusPrintLabelsShipmentOrders
{
    public partial class PrintForm : Form
    {
        public Button Button;
        public int Number { get; set; }
        public bool buttonOK_Clicked = false;
        public PrintForm(int number)
        {
            InitializeComponent();
            numericUpDown1.Value = number;
            Button = buttonOK;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Number = (int)(sender as NumericUpDown).Value;
            
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            buttonOK_Clicked = true;
        }
    }
}
