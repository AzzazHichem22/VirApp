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
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Data;

namespace WpfApplication3
{
    /// <summary>
    /// Interaction logic for LOGIN.xaml
    /// </summary>
    public partial class LOGIN : MahApps.Metro.Controls.MetroWindow
    {
        public LOGIN()
        {
            InitializeComponent();
        }


        static string conString = "Data Source=THEPUNISHER;Initial Catalog=OeuvresSociales2;Integrated Security=True";
        static SqlConnection connect = new SqlConnection(conString);
        string Query;
        static SqlCommand command = new SqlCommand();

        public string getExcuteScalar(string Query)
        {
            SqlCommand cmd = new SqlCommand(Query, connect);
            string result = "";

            if (connect.State == ConnectionState.Closed) { connect.Open(); }
            //MessageBox.Show(Query);
            if (cmd.ExecuteScalar() != null) result = cmd.ExecuteScalar().ToString();
            else result = null;
            if (connect.State == ConnectionState.Open) { connect.Close(); }

            return result;
        }



        private void Login_(object sender, RoutedEventArgs e)
        {
            MainWindow w = new MainWindow();
            string user = "";
            string mdp = "";

            if ((string.IsNullOrEmpty(NomUtilisateur.Text)) || (string.IsNullOrEmpty(MotPasse.Password)) || (NomUtilisateur.Text == "Nom d'utilisateur") || (MotPasse.Password == "Mot de passe"))
            {
                MessageBox.Show("Veuillez completer les champs de connection");
            }
            else
            {

                Query = "SELECT Login FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "'";

                
            user = getExcuteScalar(Query);

            Query = "SELECT MotPasse FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "' AND MotPasse='" + MotPasse.Password + "'";

            mdp = getExcuteScalar(Query);

           
                if (NomUtilisateur.Text == user && MotPasse.Password == mdp)
                {
                    Query = "SELECT CodeUser FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "' AND MotPasse='" + MotPasse.Password + "'";

                    MainWindow.CodeUser = Int32.Parse(getExcuteScalar(Query));

                    Query = "SELECT droit FROM dbo.utilisateur WHERE CodeUser=" + MainWindow.CodeUser;
                    
                    MainWindow.Droit = getExcuteScalar(Query);
                    WorkSpace WS = new WorkSpace();

                    

                    w.Content = WS;
                    w.WindowState = WindowState.Maximized;
                    w.Show();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Nom d'utilisateur ou mot de passe erroné.");
                }

            }


        }

        private void NomUtilisateur_GotFocus(object sender, RoutedEventArgs e)
        {
            NomUtilisateur.Text = "";
        }

        private void MotPasse_GotFocus(object sender, RoutedEventArgs e)
        {
            MotPasse.Password = "";
        }

        private void MotPasse_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow w = new MainWindow();
                string user = "";
                string mdp = "";

                if ((string.IsNullOrEmpty(NomUtilisateur.Text)) || (string.IsNullOrEmpty(MotPasse.Password)) || (NomUtilisateur.Text == "Nom d'utilisateur") || (MotPasse.Password == "Mot de passe"))
                {
                    MessageBox.Show("Veuillez completer les champs de connection");
                }
                else
                {

                    Query = "SELECT Login FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "'";


                    user = getExcuteScalar(Query);

                    Query = "SELECT MotPasse FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "' AND MotPasse='" + MotPasse.Password + "'";

                    mdp = getExcuteScalar(Query);


                    if (NomUtilisateur.Text == user && MotPasse.Password == mdp)
                    {
                        Query = "SELECT CodeUser FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "' AND MotPasse='" + MotPasse.Password + "'";

                        MainWindow.CodeUser = Int32.Parse(getExcuteScalar(Query));

                        Query = "SELECT droit FROM dbo.utilisateur WHERE CodeUser=" + MainWindow.CodeUser;

                        MainWindow.Droit = getExcuteScalar(Query);
                        WorkSpace WS = new WorkSpace();



                        w.Content = WS;
                        w.WindowState = WindowState.Maximized;
                        w.Show();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Nom d'utilisateur ou mot de passe erroné.");
                    }

                }
            }
        }




        private void NomUtilisateur_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow w = new MainWindow();
                string user = "";
                string mdp = "";

                if ((string.IsNullOrEmpty(NomUtilisateur.Text)) || (string.IsNullOrEmpty(MotPasse.Password)) || (NomUtilisateur.Text == "Nom d'utilisateur") || (MotPasse.Password == "Mot de passe"))
                {
                    MessageBox.Show("Veuillez completer les champs de connection");
                }
                else
                {

                    Query = "SELECT Login FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "'";


                    user = getExcuteScalar(Query);

                    Query = "SELECT MotPasse FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "' AND MotPasse='" + MotPasse.Password + "'";

                    mdp = getExcuteScalar(Query);


                    if (NomUtilisateur.Text == user && MotPasse.Password == mdp)
                    {
                        Query = "SELECT CodeUser FROM dbo.utilisateur WHERE Login='" + NomUtilisateur.Text + "' AND MotPasse='" + MotPasse.Password + "'";

                        MainWindow.CodeUser = Int32.Parse(getExcuteScalar(Query));

                        Query = "SELECT droit FROM dbo.utilisateur WHERE CodeUser=" + MainWindow.CodeUser;

                        MainWindow.Droit = getExcuteScalar(Query);
                        WorkSpace WS = new WorkSpace();



                        w.Content = WS;
                        w.WindowState = WindowState.Maximized;
                        w.Show();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Nom d'utilisateur ou mot de passe erroné.");
                    }

                }
            }
        }
    }
    

    }

