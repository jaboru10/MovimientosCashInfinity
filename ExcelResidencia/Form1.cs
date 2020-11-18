using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace ExcelResidencia
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        private const string EXCEL = "select ges301.tfechahora as 'FECHA'," +
        "ges301.ntipo as 'OPERACION',"+
        "empleado.cempleado as 'EMPLEADO'," +
        "ges301.icaja as 'CAJA'," +
        "ges301.bimporte as 'IMPORTE', " +
        "ges301.mticket as 'TICKET', " +
        "ges309.cconcepto as 'CONCEPTO 309', " +
        "ges310.cconcepto as 'CONCEPTO310', "+
        "ges305.cdescripcion as 'descricion entrega a cuenta' "+
        "from ges301 " +
        "LEFT JOIN ges065 on ges065.idoperacion=ges301.idoperacion " +
        "LEFT JOIN empleado on empleado.id=ges301.idempleado " +
        "LEFT JOIN ges309 on ges309.idoperacion=ges301.idoperacion "+
        "LEFT JOIN ges310 on ges310.idoperacion= ges301.idoperacion " +
        "LEFT JOIN ges305 on ges305.idoperacion=ges301.idoperacion " +
        "where (ges301.ntipo=6 or ges301.ntipo=7 or ges301.ntipo=9) and ges065.idoperacion is null ";



        private const string EXCEL2 = "select ges301.tfechahora as 'FECHA'," +
       "ges301.ntipo as 'OPERACION'," +
       "empleado.cempleado as 'EMPLEADO'," +
       "ges301.icaja as 'CAJA'," +
       "ges301.bimporte as 'IMPORTE', " +
       "ges301.mticket as 'TICKET' " +
       "from ges301 " +
       "LEFT JOIN ges065 on ges065.idoperacion=ges301.idoperacion " +
       "LEFT JOIN empleado on empleado.id=ges301.idempleado " +
       "where (ges301.ntipo=1 or ges301.ntipo=2 or ges301.ntipo=3 or ges301.ntipo=6 or ges301.ntipo=7 or ges301.ntipo=9) and ges065.idoperacion is null ";

        public Form1()
        {
            InitializeComponent();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {

            string promptValue = Prompt.ShowDialog("Indique un nombre para el excel", "Nombre archivo csv");
            try
            {
                if (promptValue == "")
                {
                    MessageBox.Show("No se ha indicado nombre del fichero, por favor indique nombre del fichero.");
                }
                else
                {
           

                    string SQL = EXCEL +  " and ges301.tfechahora between '" +
                        mdtDesde.Value.ToString("yyyy-MM-dd 00:00:00") + "' and '" + mdtHasta.Value.ToString("yyyy-MM-dd 23:59:59") + "'";                    

                    Farmanager.Farmanager farm = new Farmanager.Farmanager(Properties.Settings.Default.host, 3306);
                    FileInfo fi = new FileInfo(promptValue + ".csv");
                    List<object[]> resultados = farm.Select(SQL);
                    StreamWriter sw = new StreamWriter(fi.FullName, false, Encoding.UTF8);
                    sw.WriteLine("\"FECHA\";\"OPERACION\";\"EMPLEADO\";\"CAJA\";\"IMPORTE\";");
                    double importeTotal = 0;
                    foreach (object[] linea in resultados)
                    {
                        int ntipo = (int)linea[1];
                        string valor = "";

                        if (ntipo == 9)
                        {
                            linea[1] = "RETIRADA";

                            string ticket = linea[5].ToString();
                            string[] valores = ticket.Split(':');
                            string cadenaok = valores[3].ToString();
                            string[] arrayfinal = cadenaok.Split(' ');
                            string cadenafinal = arrayfinal[9].ToString();

                            if (cadenafinal.Length > 5)
                            {
                                if (cadenafinal.Contains("Divi"))
                                {
                                    cadenafinal = arrayfinal[7].ToString();
                                }
                                
                                    string str = cadenafinal.Remove(cadenafinal.Length - 2);
                                    linea[4] = "-" + str;
                                
                            }
                            else
                            {
                                linea[5] = cadenafinal;
                            }

                        }
                        else
                        {
                            if (linea[7] != System.DBNull.Value)
                            {
                                if (linea[7] != "")
                                {
                                    linea[1] = linea[7];
                                    linea[5] = "";
                                    linea[4] = "-" + linea[4].ToString();
                                }
                                else
                                {
                                    linea[1] = "PAGO CAJA";
                                    linea[5] = "";
                                }
                            }
                            if (linea[6] != System.DBNull.Value)
                            {
                                linea[1] = linea[6];
                                linea[5] = "";
                            }
                            if (linea[8] != System.DBNull.Value)
                            {
                                linea[1] = linea[8];
                                linea[5] = "";
                            }


                            
                            

                            if (linea[1] == System.DBNull.Value)
                            {
                          
                                if (linea[1].Equals(7))
                                {
                                    linea[1] = "PAGO CAJA";
                                    linea[5] = "";
                                }


                            }


                     
                        }
                        linea[5] = ntipo;
                        linea[6] = "";
                        linea[7] = "";
                        linea[8] = "";
                        
                        
                        try { importeTotal += Convert.ToDouble(linea[4]); }
                        catch {
                            importeTotal = 0;
                        }
                       
                        foreach (var d in linea)
                        {
                         
                            valor = valor + "\"" + d + "\";";
                        }
                        sw.WriteLine(valor);
                    }

                    sw.Close();
                    if (fi.Exists)
                    {
                        Process.Start(fi.FullName);
                    }
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if(DateTime.Today.Month == 1)
            {
                mdtDesde.Value = new DateTime(DateTime.Today.Year - 1, 12, 1);
            }
            else
            {
                mdtDesde.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month - 1, 1);
            }

            mdtHasta.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            Configuracion conf = new Configuracion();
            if (conf.ShowDialog(this) == DialogResult.OK)
            {

                conf.Dispose();
            }
        }
        public static class Prompt
        {
            public static string ShowDialog(string text, string caption)
            {
                Form prompt = new Form()
                {
                    Width = 500,
                    Height = 150,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = caption,
                    StartPosition = FormStartPosition.CenterScreen
                };
                Label textLabel = new Label() { Left = 50, Top = 20, Text = text };
                TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
                Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
                confirmation.Click += (sender, e) => { prompt.Close(); };
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.Controls.Add(textLabel);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }
        }
    }
}


/*
 * if (linea[1].Equals(1))
{
    linea[1] = "VENTA FINALIZADA";
    linea[5] = "";

}
else if (linea[7] == null)
{

}
if(linea[6] != System.DBNull.Value)
{
    linea[5] = "";
}
if (linea[1].Equals(2))
{
    linea[1] = "ENTREGA A CUENTA DEL CLIENTE";
    linea[5] = "";

}
else if (linea[1].Equals(3))
{
    linea[1] = "COBRO LINEA A CLIENTE";
    linea[5] = "";
}
else if (linea[1].Equals(6))
{
    linea[1] = "COBRO CAJA";
    linea[5] = "";
}
else if (linea[1].Equals(7))
{
    linea[1] = "PAGO CAJA";
    linea[5] = "";
}
else if (linea[1].Equals(9))
{
    linea[1] = "RETIRADA";

    string ticket = linea[5].ToString();
    string[] valores = ticket.Split(':');
    string cadenaok = valores[3].ToString();
    string[] arrayfinal = cadenaok.Split(' ');
    string cadenafinal = arrayfinal[9].ToString();

    if (cadenafinal.Length > 5)
    {
        string str = cadenafinal.Remove(cadenafinal.Length-2);
        linea[5] = str;
    }
    else
    {
        linea[5] = cadenafinal;
    }
 */