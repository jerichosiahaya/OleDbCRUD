using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OleDbCRUD
{
    public partial class Form2 : Form
    {
        //ctor
        private readonly DataTable _dt;
        public Form2(DataTable dataTable)
        {
            InitializeComponent();
            _dt = dataTable;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = _dt;
        }
    }
}
