using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MessageBoxAddin.Forms
{
    public partial class MyCustomForm : Form
    {
        public string TextBoxContents => textBox1.Text.Trim();

        public MyCustomForm()
        {
            InitializeComponent();
        }
    }
}
