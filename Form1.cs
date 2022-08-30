using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
namespace Proyecto_Beca
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataView ImportarDatos(string nombreArchivo)
        {
            string conexion = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0;'", nombreArchivo);

            OleDbConnection conector = new OleDbConnection(conexion);

            conector.Open();

            OleDbCommand consulta = new OleDbCommand("select * from [Hoja 1$]", conector);

            OleDbDataAdapter adaptador = new OleDbDataAdapter
            {
                SelectCommand = consulta
            };
            DataSet ds = new DataSet();

            adaptador.Fill(ds);

            conector.Close();

            return ds.Tables[0].DefaultView;
        }
        private void button3_Click(object sender, EventArgs e) {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel | *.xls;*.xlsx;",
                Title = "Seleccionar Archivo"
            };

            if (openFileDialog.ShowDialog()==DialogResult.OK)
            {
                dataGridView1.DataSource = ImportarDatos(openFileDialog.FileName);
            }

            //Proceso para convertir a mayuscula
            //Variable
            string[] nombre = new string[3];
            int i;

            i = 0;
            while (i < nombre.Length)
            {
                nombre[i] = dataGridView1.Rows[i].Cells[1].Value.ToString();

                string[] nombreSeparado = nombre[i].Split(new char[] {' ', ','});

                string apellidoSep = nombreSeparado[0];
                string nombreSep = nombreSeparado[1];

                string ape = char.ToUpper(apellidoSep[0]) + apellidoSep.Substring(1); ;
                string nom = char.ToUpper(nombreSep[0]) + nombreSep.Substring(1);

                string ApellidoYnombre = ape + nom;

                dataGridView1.Rows[i].Cells[1].Value = ApellidoYnombre;

                i++;
            }
        }

    }
}
