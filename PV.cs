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
    class PV
    {
        private int code_PV;
        private DateTime date_PV;
 
        public PV(int code)
        {
            this.code_PV = code;
            
        }

        public void Add_PV()
        {
            try
            {
                BD con = new BD();
                con.seConnecter();
                String requette = "INSERT INTO PV ( CodePV, DateCreatPV, CodeUser) VALUES('" + code_PV + "','" + DateTime.Today + "',01)";
                SqlCommand cmd = new SqlCommand(requette, con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show(" réussite de l ' addition ");
                con.seDeconnecter();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(" Oups !Ce PV existe deja ! \n" + ex.Message);
            }
        }
    }
}
