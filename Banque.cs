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
    class Banque
    {
        private int code_banque;
        private string designation_banque;
        private string montant_banque;
        public Banque(string name, string  montant)
        {

            this.designation_banque = name;
            this.montant_banque = montant;
        }
        public void Add_Banque()
        {
            try
            {
                BD con = new BD();
                con.seConnecter();
                int i = con.nbLigne("Banque", " WHERE  DésignationBanque = '" + designation_banque + "' ");
                if (i < 1)
                {
                    int code = con.nbLigne("Banque", " ") + 1;
                    String requette = "INSERT INTO Banque (CodeBanque, DésignationBanque, AdrBaque,DateCreatType, CodeUser ) VALUES('" + code + "','" + designation_banque + "','" + montant_banque + "','" + DateTime.Today + "',01)";
                    SqlCommand cmd = new SqlCommand(requette, con.connextion());
                    cmd.ExecuteNonQuery();
                    MessageBox.Show(" réussite de l ' addition ");
                    con.seDeconnecter();
                }
                else
                {
                    MessageBox.Show(" Oups!! Cette Banque  existe deja ! \n");
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(" Une erreur a éte produite ! \n" + ex.Message);
            }
        }
    }
}
