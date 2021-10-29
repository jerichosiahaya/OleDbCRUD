using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OleDbCRUD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\praktikum-visprog\\OleDbCRUD\\db.accdb;Persist Security Info=False;";

        // insert
        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection vConnect = new OleDbConnection(connection); // connection
            string qInsert = "insert into dataMahasiswa (nim, nama) values (@nim, @nama)"; // query
            OleDbCommand vInsert = new OleDbCommand(qInsert, vConnect);
            vInsert.Parameters.AddWithValue("@nim", Convert.ToInt32(textBox1.Text)); // isi parameternya
            vInsert.Parameters.AddWithValue("@nama", textBox2.Text);
            try
            {
                vConnect.Open();
                OleDbDataReader vReaderInsert = vInsert.ExecuteReader();
                MessageBox.Show("Data berhasil ditambah");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                vConnect.Close();
            }
        }

        // update
        private void button2_Click(object sender, EventArgs e)
        {
            OleDbConnection vConnect = new OleDbConnection(connection); // connection
            string qUpdate = "update dataMahasiswa set nama = @nama where nim = @nim"; // query
            OleDbCommand vUpdate = new OleDbCommand(qUpdate, vConnect);
            vUpdate.Parameters.AddWithValue("@nama", textBox2.Text); // isi parameter sesuai urutan di query string
            vUpdate.Parameters.AddWithValue("@nim", Convert.ToInt32(textBox1.Text)); // isi parameternya
            try
            {
                vConnect.Open();
                OleDbDataReader vReaderInsert = vUpdate.ExecuteReader();
                MessageBox.Show("Data berhasil diupdate");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                vConnect.Close();
            }
        }

        // delete
        private void button3_Click(object sender, EventArgs e)
        {
            OleDbConnection vConnect = new OleDbConnection(connection); // connection
            string qDelete = "delete from dataMahasiswa where nim = @nim"; // query
            OleDbCommand vDelete = new OleDbCommand(qDelete, vConnect);
            //vDelete.Parameters.AddWithValue("@nama", textBox2.Text); // isi parameter sesuai urutan di query string
            vDelete.Parameters.AddWithValue("@nim", Convert.ToInt32(textBox1.Text)); // isi parameternya
            try
            {
                vConnect.Open();
                OleDbDataReader vReaderInsert = vDelete.ExecuteReader();
                MessageBox.Show("Data berhasil dihapus");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                vConnect.Close();
            }
        }

        // retrieve
        private void button4_Click(object sender, EventArgs e)
        {
            OleDbConnection vConnect = new OleDbConnection(connection); // connection
            string qRetrieve = "select * from dataMahasiswa where nim = @nim"; // query
            OleDbCommand vRetrieve = new OleDbCommand(qRetrieve, vConnect);
            //vDelete.Parameters.AddWithValue("@nama", textBox2.Text); // isi parameter sesuai urutan di query string
            vRetrieve.Parameters.AddWithValue("@nim", Convert.ToInt32(textBox1.Text)); // isi parameternya
            try
            {
                

                var dt = new DataTable(); // buat datatable
                vConnect.Open();
                OleDbDataReader vReaderInsert = vRetrieve.ExecuteReader();
                dt.Load(vReaderInsert); //datatable diload dari obj oledbdatareader

                Form2 vFrom = new Form2(dt);
                vFrom.Show();
                //MessageBox.Show("Data berhasil dihapus");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                vConnect.Close();
            }
        }
    }
}
