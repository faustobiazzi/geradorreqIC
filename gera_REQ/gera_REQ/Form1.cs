using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace gera_REQ
{
    public partial class insert_Dados : Form
    {
        

        public insert_Dados()
        {
            InitializeComponent();
            Hora_Cepol.CustomFormat = "hh'H'mm";
            datahoraPericia.CustomFormat = "dd/MM/yy - hh'H'mm";
            hora_Resposta.CustomFormat = "hh'H'mm";


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
       
        private void BTN_GERARREQ_Click(object sender, EventArgs e)
        {

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filePath = saveFileDialog1.FileName;
                string respostaDP = ",";
                string respostaCEPOL = ",";



                string v1 = TXT_Natureza.Text + ",";
                string v2 = cmbOrigem.Text + ",";
                string v3 = TXT_BOPM.Text+",";
                string v4 = data_Acionamento.Value.ToString() + ",";
                string v5 = Hora_Cepol.Value.ToString() + ",";
                if (check_DP.Checked)
                { 
                    respostaDP = "X,";
                }
                if (checkCEPOL.Checked)
                    respostaCEPOL = "X,";
                string v6 = hora_Resposta.Value.ToString() + ",";
                string V7 = datahoraPericia.Value.ToString() + ",";
                string v8 = TXT_localFato.Text + ",";
                string v9 = txt_REF.Text + ",";
                string v10;
                if (preservadoSIM.Checked) { 
                    v10 = "preservado,";
                }
                else
                {
                    v10 = "não preservado,";
                }



                StringBuilder csvcontent = new StringBuilder();
                csvcontent.AppendLine("natureza,origem,bopm,data de acionamento, hora cepol, resposta DP,Resposta Cepol, " +
                    "hora da resposta, pericia no local,local do fato, referência, status de preservação,");
                csvcontent.AppendLine(v1 + v2 + v3 + v4 + v5+respostaDP+respostaCEPOL+v6+V7+v8+v9+v10);
                //File.AppendAllText(filePath, csvcontent.ToString(), Encoding.GetEncoding(28591));
                
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    File.AppendAllText(filePath, csvcontent.ToString(), Encoding.GetEncoding(28591));
                    MessageBox.Show("complete");
                }
                else
                {
                    File.AppendAllText(filePath, csvcontent.ToString(), Encoding.GetEncoding(28591));
                    
                }
            }
        }

        private void BTN_ABRIR_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            string file = "";
        
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                
                OleDbConnection cnn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + "; Extended Properties=Excel 12.0;");

                OleDbCommand oconn = new OleDbCommand("SELECT * from [Plan1$]", cnn);
                cnn.Open();
                OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                adp.Fill(dt);

               

            }
        }

        private void btn_Exportar_Click(object sender, EventArgs e)
        {

        }
    }    
}
