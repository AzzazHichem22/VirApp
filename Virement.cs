using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Drawing;
using System.Data.SqlClient;
using System.Data;

namespace WpfApplication3
{
    class Virement
    {
        private int code_Virement;
        private DateTime date_Virement;
        private int code_PV;
        public Virement(int code, DateTime date, int code_v)
        {
            this.code_Virement = code;
            this.date_Virement = date;
            this.code_PV = code_v;
        }

        public void Add_Virement()
        {
            try
            {
                BD con = new BD();
                con.seConnecter();
                String requette = "INSERT INTO Virement (CodeVir, DateVir, pv_codepv) VALUES('" + code_Virement + "','" + date_Virement + "','" + code_PV + "')";
                SqlCommand cmd = new SqlCommand(requette, con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show(" réussite de l ' addition ");
                con.seDeconnecter();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(" Oups !Ce fonctionnaire existe deja ! \n" + ex.Message);
            }
        }
    }
}
