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
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data.Common;
using System.Data;
using System.IO;
using Microsoft.Win32;

namespace WpfApplication3
{

    public partial class WorkSpace : UserControl
    {
        public WorkSpace()
        {
            InitializeComponent();
            Upload();
            Upload_traitement();

        }

        


        // Partie visuel

        public enum choix
        {
            Mise_a_jour,
            Demande,
            Virement,
            Statistiques
        }


        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (((ListBox)sender).SelectedIndex)
            {
                case 0:
                    accueil.Visibility = Visibility.Visible;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 1:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Visible;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 2:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Visible;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 3:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Visible;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 4:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Visible;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 5:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Visible;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 6:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Visible;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;

                case 7:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Visible;

                    break;

                case 8:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Visible;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
            }
        }

        // Attributs de la classe

        static string conString = "Data Source=THEPUNISHER;Initial Catalog=OeuvresSociales2;Integrated Security=True";
        static SqlConnection connect = new SqlConnection(conString);
        string Query;
        static SqlCommand command = new SqlCommand();

        public static int CodeUser = 1;

        static string lastKeyQuery = "SELECT NumDem FROM DemandePrime";
        static string lastKey = getLastKey(lastKeyQuery);


        // Methodes de la classe 

        public string verifNumDemaAndTypePrime(string NumDemFormulaire)
        {
            Query = "SELECT CodePrime FROM DemandePrime ";
            Query += " WHERE NumDem=" + NumDemFormulaire;

            string CodePrime = getExcuteScalar(Query);

            Query = "SELECT DésignationPrime FROM TypePrime ";
            Query += " WHERE CodePrime=" + CodePrime;

            string DésignationPrime = getExcuteScalar(Query);

            if (DésignationPrime == "Don") return "Don";
            else if (DésignationPrime == "Décès Employée") return "Décès Employée";
            else if (DésignationPrime == "") return "";
            else return "Autres";



        }

        public string getExcuteScalar(string Query)
        {
            SqlCommand cmd = new SqlCommand(Query, connect);
            string result = "";

            if (connect.State == ConnectionState.Closed) { connect.Open(); }
            //MessageBox.Show(Query);
            result = cmd.ExecuteScalar().ToString();

            if (connect.State == ConnectionState.Open) { connect.Close(); }

            return result;
        }

        public class ConvertisseurChiffresLettres
        {

            public string convertion(double chiffre)
            {
                int centaine, dizaine, unite, reste, y;
                bool dix = false;
                bool soixanteDix = false;
                string lettre = "";

                reste = (int)chiffre / 1;

                for (int i = 1000000000; i >= 1; i /= 1000)
                {
                    y = reste / i;
                    if (y != 0)
                    {
                        centaine = y / 100;
                        dizaine = (y - centaine * 100) / 10;
                        unite = y - (centaine * 100) - (dizaine * 10);
                        switch (centaine)
                        {
                            case 0:
                                break;
                            case 1:
                                lettre += "cent ";
                                break;
                            case 2:
                                if ((dizaine == 0) && (unite == 0)) lettre += "deux cents ";
                                else lettre += "deux-cent ";
                                break;
                            case 3:
                                if ((dizaine == 0) && (unite == 0)) lettre += "trois cents ";
                                else lettre += "trois-cent ";
                                break;
                            case 4:
                                if ((dizaine == 0) && (unite == 0)) lettre += "quatre cents ";
                                else lettre += "quatre-cent ";
                                break;
                            case 5:
                                if ((dizaine == 0) && (unite == 0)) lettre += "cinq cents ";
                                else lettre += "cinq-cent ";
                                break;
                            case 6:
                                if ((dizaine == 0) && (unite == 0)) lettre += "six cents ";
                                else lettre += "six-cent ";
                                break;
                            case 7:
                                if ((dizaine == 0) && (unite == 0)) lettre += "sept cents ";
                                else lettre += "sept-cent ";
                                break;
                            case 8:
                                if ((dizaine == 0) && (unite == 0)) lettre += "huit cents ";
                                else lettre += "huit-cent ";
                                break;
                            case 9:
                                if ((dizaine == 0) && (unite == 0)) lettre += "neuf cents ";
                                else lettre += "neuf-cent ";
                                break;
                        }// La fin du cas " centaine "

                        switch (dizaine)
                        {
                            case 0:
                                break;
                            case 1:
                                dix = true;
                                break;
                            case 2:
                                lettre += "vingt ";
                                break;
                            case 3:
                                lettre += "trente ";
                                break;
                            case 4:
                                lettre += "quarante ";
                                break;
                            case 5:
                                lettre += "cinquante ";
                                break;
                            case 6:
                                lettre += "soixante ";
                                break;
                            case 7:
                                dix = true;
                                soixanteDix = true;
                                lettre += "soixante ";
                                break;
                            case 8:
                                lettre += "quatre-vingt ";
                                break;
                            case 9:
                                dix = true;
                                lettre += "quatre-vingt ";
                                break;
                        } // La fin du cas " dizaine "

                        switch (unite)
                        {
                            case 0:
                                if (dix) lettre += "dix ";
                                break;
                            case 1:
                                if (soixanteDix) lettre += "et onze ";
                                else
                                    if (dix) lettre += "onze ";
                                else if ((dizaine != 1 && dizaine != 0)) lettre += "et un ";
                                else lettre += "un ";
                                break;
                            case 2:
                                if (dix) lettre += "douze ";
                                else lettre += "quatre ";
                                break;
                            case 5:
                                if (dix) lettre += "quinze ";
                                else lettre += "cinq ";
                                break;
                            case 6:
                                if (dix) lettre += "seize ";
                                else lettre += "six ";
                                break;
                            case 7:
                                if (dix) lettre += "dix-sept ";
                                else lettre += "sept ";
                                break;
                            case 8:
                                if (dix) lettre += "dix-huit ";
                                else lettre += "huit ";
                                break;
                            case 9:
                                if (dix) lettre += "dix-neuf ";
                                else lettre += "neuf ";
                                break;
                        } // La fin du cas " unite "

                        switch (i)
                        {
                            case 1000000000:
                                if (y > 1) lettre += "milliards ";
                                else lettre += "milliard ";
                                break;
                            case 1000000:
                                if (y > 1) lettre += "millions ";
                                else lettre += "million ";
                                break;
                            case 1000:
                                lettre += "mille ";
                                break;
                        }
                    } // la fin de la condition if ( y!= 0 )
                    reste -= y * i;
                    dix = false;
                    soixanteDix = false;
                } // la fin de la boucle "pour" 

                if (lettre.Length == 0) lettre += "zero";

                // pour les chiffres apres la virgule :

                Decimal chiffresDecimals;
                chiffresDecimals = (Decimal)(chiffre * 100) % 100;


                dizaine = (int)(chiffresDecimals) / 10;
                unite = (int)chiffresDecimals - (dizaine * 10);

                string lettreDecimal = "";
                switch (dizaine)
                {
                    case 0:
                        break;
                    case 1:
                        dix = true;
                        break;
                    case 2:
                        lettreDecimal += "vingt ";
                        break;
                    case 3:
                        lettreDecimal += "trente ";
                        break;
                    case 4:
                        lettreDecimal += "quarante ";
                        break;
                    case 5:
                        lettreDecimal += "cinquante ";
                        break;
                    case 6:
                        lettreDecimal += "soixante ";
                        break;
                    case 7:
                        dix = true;
                        soixanteDix = true;
                        lettreDecimal += "soixante ";
                        break;
                    case 8:
                        lettreDecimal += "quatre-vingt ";
                        break;
                    case 9:
                        dix = true;
                        lettreDecimal += "quatre-vingt ";
                        break;
                } // La fin du cas " dizaine "

                switch (unite)
                {
                    case 0:
                        if (dix) lettreDecimal += "dix ";
                        break;
                    case 1:
                        if (soixanteDix) lettreDecimal += "et onze ";
                        else
                            if (dix) lettreDecimal += "onze ";
                        else if ((dizaine != 1 && dizaine != 0)) lettreDecimal += "et un ";
                        else lettreDecimal += "un ";
                        break;
                    case 2:
                        if (dix) lettreDecimal += "douze ";
                        else lettreDecimal += "deux ";
                        break;
                    case 3:
                        if (dix) lettreDecimal += "treize ";
                        else lettreDecimal += "trois ";
                        break;
                    case 4:
                        if (dix) lettreDecimal += "quatorze ";
                        else lettreDecimal += "quatre ";
                        break;
                    case 5:
                        if (dix) lettreDecimal += "quinze ";
                        else lettreDecimal += "cinq ";
                        break;
                    case 6:
                        if (dix) lettreDecimal += "seize ";
                        else lettreDecimal += "six ";
                        break;
                    case 7:
                        if (dix) lettreDecimal += "dix-sept ";
                        else lettreDecimal += "sept ";
                        break;
                    case 8:
                        if (dix) lettreDecimal += "dix-huit ";
                        else lettreDecimal += "huit ";
                        break;
                    case 9:
                        if (dix) lettreDecimal += "dix-neuf ";
                        else lettreDecimal += "neuf ";
                        break;
                } // La fin du cas " unite "


                // Traiter le cas de " un mille " :

                if (lettre.StartsWith("un mille")) lettre = lettre.Remove(0, 3);

                /* Rajouter la devise ( Dinars ) et traitement des cas spéciaux */

                if (lettreDecimal.Equals(""))
                {
                    if (lettre.Equals("un "))
                        return lettre + "dinar";
                    else
                        return lettre + "dinars";
                }
                else if (dizaine.Equals(0) && unite.Equals(1))
                {
                    if (lettre.Equals("un "))
                        return lettre + "dinar et " + lettreDecimal + "centime";
                    else
                        return lettre + "dinars et " + lettreDecimal + "centime";
                }

                else
                    return lettre + "dinars et " + lettreDecimal + "centimes";
            }


            // Methode pour mettre la première lettre en majuscule
            public string PremiereLettreMaj(string ChaineAConvertir)
            {
                if (!(String.IsNullOrEmpty(ChaineAConvertir)))
                {
                    return ChaineAConvertir.First().ToString().ToUpper() + String.Join("", ChaineAConvertir.Skip(1));
                }
                else
                {
                    return ChaineAConvertir;
                }
            }
        }

        static public string getLastKey(string Query)
        {
            string LastKey = "";

            if (connect.State == ConnectionState.Closed) { connect.Open(); }

            SqlCommand cmd = new SqlCommand(Query, connect);

            SqlDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                LastKey = reader[0].ToString();
            }

            if (connect.State == ConnectionState.Open) { connect.Close(); }

            return LastKey;
        }



        // Partie Formulaire 


        private void Type_formulaire_Initialized(object sender, EventArgs e)
        {
            Type_formulaire.Items.Add("Décès Employée");
            Type_formulaire.Items.Add("Don");
            Type_formulaire.Items.Add("Autres");
        }


        private void remplir_formulaire_Click(object sender, RoutedEventArgs e)
        {
            if (Type_formulaire.SelectedItem == null || Num_demande_formulaire.Text == "_____")
            {
                MessageBox.Show("Veuillez choisir un type de prime et/ou donner le numero de demande !");

            }



            else
            {
                if ((string)Type_formulaire.SelectedItem == "Autres" && verifNumDemaAndTypePrime(Num_demande_formulaire.Text) == "Autres")
                {
                    var application = new Microsoft.Office.Interop.Word.Application();
                    var Formulaire = new Microsoft.Office.Interop.Word.Document();

                    Formulaire = application.Documents.Add(Template: @"C:\Users\yasmine\Documents\Formulaire.dotx");


                    foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                    {
                        if (field.Code.Text.Contains("NumDem"))
                        {
                            try
                            {
                                field.Select();
                                application.Selection.TypeText(Num_demande_formulaire.Text);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Date"))
                        {
                            field.Select();
                            application.Selection.TypeText(DateTime.Now.ToString("dd/MM/yyyy"));
                        }


                        else if (field.Code.Text.Contains("SitFam"))
                        {
                            field.Select();
                            Query = "SELECT SitFamFonct FROM dbo.Fonctionnaire ";
                            Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                            Query += " WHERE NumDem =" + Num_demande_formulaire.Text;

                            application.Selection.TypeText(getExcuteScalar(Query));

                        }

                        else if (field.Code.Text.Contains("NomPrenom"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT NomFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string NomFonct = getExcuteScalar(Query);

                                Query = "SELECT PrenFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string PrenFonct = getExcuteScalar(Query);

                                application.Selection.TypeText(NomFonct + " " + PrenFonct);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }




                        else if (field.Code.Text.Contains("TelFonct"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT TelFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                application.Selection.TypeText("0" + getExcuteScalar(Query));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Aide"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT DésignationPrime FROM dbo.TypePrime ";
                                Query += " INNER JOIN DemandePrime ON TypePrime.CodePrime = DemandePrime.CodePrime";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                application.Selection.TypeText(getExcuteScalar(Query));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Montant"))
                        {

                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                field.Select();
                                Query = "SELECT MontantPrime FROM dbo.TypePrime ";
                                Query += " INNER JOIN DemandePrime ON TypePrime.CodePrime = DemandePrime.CodePrime";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                command = new SqlCommand(Query, connect);
                                ConvertisseurChiffresLettres chiffre = new ConvertisseurChiffresLettres();
                                application.Selection.TypeText(command.ExecuteScalar().ToString() + " DA    ( " + chiffre.PremiereLettreMaj(chiffre.convertion((double)command.ExecuteScalar())) + " )");
                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }


                    Formulaire.SaveAs2(FileName: @".\ProjectTests\Demande_" + Num_demande_formulaire.Text + ".docx");

                    application.Visible = true;

                    //application.Quit();
                }



                else if ((string)Type_formulaire.SelectedItem == "Don" && verifNumDemaAndTypePrime(Num_demande_formulaire.Text) == "Don")
                {
                    var application = new Microsoft.Office.Interop.Word.Application();
                    var Formulaire = new Microsoft.Office.Interop.Word.Document();

                    Formulaire = application.Documents.Add(Template: @"C:\Users\amine\Documents\Don.docx");


                    foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                    {
                        if (field.Code.Text.Contains("NumDem"))
                        {
                            try
                            {
                                field.Select();
                                application.Selection.TypeText(Num_demande_formulaire.Text);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("date"))
                        {
                            field.Select();
                            application.Selection.TypeText(DateTime.Now.ToString("dd/MM/yyyy"));
                        }


                        else if (field.Code.Text.Contains("SitFam"))
                        {
                            field.Select();
                            Query = "SELECT SitFamFonct FROM dbo.Fonctionnaire ";
                            Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                            Query += " WHERE NumDem =" + Num_demande_formulaire.Text;

                            application.Selection.TypeText(getExcuteScalar(Query));

                        }

                        else if (field.Code.Text.Contains("NomPrenom"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT NomFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string NomFonct = getExcuteScalar(Query);

                                Query = "SELECT PrenFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string PrenFonct = getExcuteScalar(Query);

                                application.Selection.TypeText(NomFonct + " " + PrenFonct);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }




                        else if (field.Code.Text.Contains("Montant"))
                        {

                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                field.Select();
                                Query = "SELECT MontantDem FROM DemandePrime ";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                command = new SqlCommand(Query, connect);
                                ConvertisseurChiffresLettres chiffre = new ConvertisseurChiffresLettres();
                                application.Selection.TypeText(command.ExecuteScalar().ToString() + " DA    ( " + chiffre.PremiereLettreMaj(chiffre.convertion((double)command.ExecuteScalar())) + " )");
                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }


                    Formulaire.SaveAs2(FileName: @".\ProjectTests\Demande_" + Num_demande_formulaire.Text + ".docx");

                    application.Visible = true;
                }

                else if (Type_formulaire.Text == "Décès Employée" && verifNumDemaAndTypePrime(Num_demande_formulaire.Text) == "Décès Employée")
                {
                    var application = new Microsoft.Office.Interop.Word.Application();
                    var Formulaire = new Microsoft.Office.Interop.Word.Document();

                    Formulaire = application.Documents.Add(Template: @"C:\Users\amine\Documents\Décès.docx");


                    foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                    {
                        if (field.Code.Text.Contains("NumDem"))
                        {
                            try
                            {
                                field.Select();
                                application.Selection.TypeText(Num_demande_formulaire.Text);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Date"))
                        {
                            field.Select();
                            application.Selection.TypeText(DateTime.Now.ToString("dd/MM/yyyy"));
                        }


                        else if (field.Code.Text.Contains("SitFam"))
                        {
                            field.Select();
                            Query = "SELECT SitFamFonct FROM dbo.Fonctionnaire ";
                            Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                            Query += " WHERE NumDem =" + Num_demande_formulaire.Text;

                            application.Selection.TypeText(getExcuteScalar(Query));

                        }

                        else if (field.Code.Text.Contains("NomPrenom"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT NomFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string NomFonct = getExcuteScalar(Query);

                                Query = "SELECT PrenFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string PrenFonct = getExcuteScalar(Query);

                                application.Selection.TypeText(NomFonct + " " + PrenFonct);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }

                        else if (field.Code.Text.Contains("NameLastNameDem"))
                        {
                            field.Select();

                            Query = " SELECT NomParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string NomParent = getExcuteScalar(Query);

                            Query = " SELECT PrenParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string PreomParent = getExcuteScalar(Query);

                            application.Selection.TypeText(NomParent + " " + PreomParent);
                        }


                        else if (field.Code.Text.Contains("SitDem"))
                        {
                            field.Select();

                            Query = " SELECT SitFamParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string SitDem = getExcuteScalar(Query);

                            application.Selection.TypeText(SitDem);
                        }

                        else if (field.Code.Text.Contains("LienParent"))
                        {
                            field.Select();

                            Query = " SELECT LienParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string LienParent = getExcuteScalar(Query);

                            application.Selection.TypeText(LienParent);
                        }


                        else if (field.Code.Text.Contains("Montant"))
                        {

                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                field.Select();
                                Query = "SELECT MontantDem FROM DemandePrime ";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                command = new SqlCommand(Query, connect);
                                ConvertisseurChiffresLettres chiffre = new ConvertisseurChiffresLettres();
                                application.Selection.TypeText(command.ExecuteScalar().ToString() + " DA    ( " + chiffre.PremiereLettreMaj(chiffre.convertion((double)command.ExecuteScalar())) + " )");
                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }


                    Formulaire.SaveAs2(FileName: @".\ProjectTests\Demande_" + Num_demande_formulaire.Text + ".docx");

                    application.Visible = true;
                }

                else
                {
                    MessageBox.Show("Le numéro de demande et le type de prime ne conviennent pas !");

                }


            }


        }


        private void Imprimer_formulaire_Click(object sender, RoutedEventArgs e)
        {
            var application = new Microsoft.Office.Interop.Word.Application();
            var FormRempli = new Microsoft.Office.Interop.Word.Document();

            FormRempli = application.Documents.Add(Template: @"C:\Users\amine\Documents\ProjectTests\Demande_" + Num_demande_formulaire + ".docx");

            FormRempli.PrintOut();

        }





        // Partie Demande

        private void Type_demandes_Initialized(object sender, EventArgs e)
        {
            Query = "SELECT * FROM TypePrime";

            SqlCommand command = new SqlCommand(this.Query, connect);

            if (connect.State == ConnectionState.Closed) connect.Open();

            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                Type_prime_demande.Items.Add((string)reader["DésignationPrime"]);
            }

            if (connect.State == ConnectionState.Open) connect.Close();


        }

        private void Type_prime_demande_DropDownClosed(object sender, EventArgs e)
        {
            if ((Type_prime_demande.Text == "Décès Parent") || (Type_prime_demande.Text == "Décès Employée"))
            {
                Tab_Décès.Visibility = Visibility.Visible;
                TabControl_Demande.SelectedItem = Tab_Décès;
            }

            if (Type_prime_demande.Text == "Don")
            {
                Tab_Don.Visibility = Visibility.Visible;
                TabControl_Demande.SelectedItem = Tab_Don;
            }
        }

        private void Num_demande_formulaire_Initialized(object sender, EventArgs e)
        {
            //Num_demande_formulaire.Text = lastKey;
            Num_demande_formulaire.Value = lastKey;

            if (Num_demande_formulaire.Value.Length == 3) Num_demande_formulaire.Value = "00" + Num_demande_formulaire.Value;
            else if (Num_demande_formulaire.Value.Length == 4) Num_demande_formulaire.Value = "0" + Num_demande_formulaire.Value;

        }

        private void Type_deces_Initialized(object sender, EventArgs e)
        {
            Type_deces.Items.Add("Parent");
            Type_deces.Items.Add("Employée");
        }

        private void Type_deces_DropDownClosed(object sender, EventArgs e)
        {
            if (Type_deces.Text == "Parent")
            {
                // Afficher ceux du Parent
                Date_event_deces.Visibility = Visibility.Visible;
                date_evenment_deces.Visibility = Visibility.Visible;
                num_demande_deces.Visibility = Visibility.Visible;
                Num_demande_deces.Visibility = Visibility.Visible;
                nom_fonct_demande_deces.Visibility = Visibility.Visible;
                Nom_fonct_demande_deces.Visibility = Visibility.Visible;
                prenom_fonct_demande_deces.Visibility = Visibility.Visible;
                Prenom_fonct_demande_deces.Visibility = Visibility.Visible;
                Date_de_demande_deces.Visibility = Visibility.Visible;
                date_demande_deces.Visibility = Visibility.Visible;
                Ajout_demande_deces_parent.Visibility = Visibility.Visible;

                // Cacher Employee
                nom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                Nom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                prenom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                Prenom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                nom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                Nom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                prenom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                Prenom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                sit_fam_demandeur_deces_employee.Visibility = Visibility.Hidden;
                Sit_fam_demandeur_deces_employee.Visibility = Visibility.Hidden;
                lien_parenté.Visibility = Visibility.Hidden;
                Lien_parenté.Visibility = Visibility.Hidden;
                Date_event_deces_employee.Visibility = Visibility.Hidden;
                date_evenment_deces_employee.Visibility = Visibility.Hidden;
                num_demande_deces_employee.Visibility = Visibility.Hidden;
                Num_demande_deces_employee.Visibility = Visibility.Hidden;
                Date_de_demande_deces_employee.Visibility = Visibility.Hidden;
                date_demande_deces_employee.Visibility = Visibility.Hidden;
                Ajout_demande_deces_employee.Visibility = Visibility.Hidden;

            }

            if (Type_deces.Text == "Employée")
            {
                // Afficher ceux de Employee
                nom_fonct_demande_deces_employee.Visibility = Visibility.Visible;
                Nom_fonct_demande_deces_employee.Visibility = Visibility.Visible;
                prenom_fonct_demande_deces_employee.Visibility = Visibility.Visible;
                Prenom_fonct_demande_deces_employee.Visibility = Visibility.Visible;
                nom_demandeur_deces_employee.Visibility = Visibility.Visible;
                Nom_demandeur_deces_employee.Visibility = Visibility.Visible;
                prenom_demandeur_deces_employee.Visibility = Visibility.Visible;
                Prenom_demandeur_deces_employee.Visibility = Visibility.Visible;
                sit_fam_demandeur_deces_employee.Visibility = Visibility.Visible;
                Sit_fam_demandeur_deces_employee.Visibility = Visibility.Visible;
                lien_parenté.Visibility = Visibility.Visible;
                Lien_parenté.Visibility = Visibility.Visible;
                Date_event_deces_employee.Visibility = Visibility.Visible;
                date_evenment_deces_employee.Visibility = Visibility.Visible;
                num_demande_deces_employee.Visibility = Visibility.Visible;
                Num_demande_deces_employee.Visibility = Visibility.Visible;
                Date_de_demande_deces_employee.Visibility = Visibility.Visible;
                date_demande_deces_employee.Visibility = Visibility.Visible;
                Ajout_demande_deces_employee.Visibility = Visibility.Visible;


                // Cacher ceux du Parent 
                Date_event_deces.Visibility = Visibility.Hidden;
                date_evenment_deces.Visibility = Visibility.Hidden;
                num_demande_deces.Visibility = Visibility.Hidden;
                Num_demande_deces.Visibility = Visibility.Hidden;
                nom_fonct_demande_deces.Visibility = Visibility.Hidden;
                Nom_fonct_demande_deces.Visibility = Visibility.Hidden;
                prenom_fonct_demande_deces.Visibility = Visibility.Hidden;
                Prenom_fonct_demande_deces.Visibility = Visibility.Hidden;
                Date_de_demande_deces.Visibility = Visibility.Hidden;
                date_demande_deces.Visibility = Visibility.Hidden;
                Ajout_demande_deces_parent.Visibility = Visibility.Hidden;
            }
        }

        private void Sit_fam_demandeur_deces_employee_Initialized(object sender, EventArgs e)
        {
            Sit_fam_demandeur_deces_employee.Items.Add("Mr");
            Sit_fam_demandeur_deces_employee.Items.Add("Mme");
            Sit_fam_demandeur_deces_employee.Items.Add("Melle");
        }

        private void Lien_parenté_Initialized(object sender, EventArgs e)
        {
            Lien_parenté.Items.Add("Père");
            Lien_parenté.Items.Add("Mère");
            Lien_parenté.Items.Add("Frère");
            Lien_parenté.Items.Add("Soeur");
            Lien_parenté.Items.Add("Fils");
            Lien_parenté.Items.Add("Fille");
        }

        private void Num_demande_Initialized(object sender, EventArgs e)
        {
            Num_demande.IsEnabled = false;

            //Query = "SELECT NumDem FROM DemandePrime";
            //string lastKey = getLastKey(Query);


            if (lastKey == "")
            {
                Num_demande.Text = "001" + DateTime.Today.ToString("yy");
            }


            else
            {
                if (Int32.Parse(lastKey) % 100 == Int32.Parse(DateTime.Today.ToString("yy")))
                {
                    Num_demande.Text = ((Int32.Parse(lastKey) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy"));
                    if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                    else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;
                }

                else if (Int32.Parse(lastKey) % 100 < Int32.Parse(DateTime.Today.ToString("yy")))
                {
                    Num_demande.Text = "001" + DateTime.Today.ToString("yy");
                }
            }
            
            

        }

        private void Ajout_demande_Click(object sender, RoutedEventArgs e)
        {
            if ((Type_prime_demande.Text == "Décès Parent") || (Type_prime_demande.Text == "Décès Employée"))
            {
                TabControl_Demande.SelectedItem = Tab_Décès;
            }

            else if (Type_prime_demande.Text == "Don")
            {
                TabControl_Demande.SelectedItem = Tab_Don;
            }

            else
            {

                Query = "SELECT Matricule FROM Fonctionnaire ";
                Query += " WHERE NomFonct ='" + Nom_fonct_demande.Text + "' AND PrenFonct='" + Prenom_fonct_demande.Text + "'";

                string Matricule = getExcuteScalar(Query);

                Query = "SELECT CodePrime FROM TypePrime ";
                Query += " WHERE DésignationPrime = '" + Type_prime_demande.Text + "'";

                string CodePrime = getExcuteScalar(Query);

                Query = "SELECT MontantPrime FROM TypePrime";
                Query += " WHERE DésignationPrime='" + Type_prime_demande.Text + "'";

                string Montant = getExcuteScalar(Query);

                Query = "SELECT CompteFonct FROM Fonctionnaire";
                Query += " WHERE Matricule=" + Matricule;

                string CompteFonct = getExcuteScalar(Query);

                string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

                string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



                Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser)";
                Query += " VALUES (" + Num_demande.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + CodeUser + ")";


                try
                {
                    if (connect.State == ConnectionState.Closed) connect.Open();
                    command = new SqlCommand(Query, connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("La demande a été ajouté !");


                    //int partieA = Int32.Parse(lastKey) / 100;
                    //partieA = partieA+1;
                    //MessageBox.Show((partieA).ToString());

                    //lastKey = (Int32.Parse(lastKey) / 100 + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy"));

                    //MessageBox.Show(lastKey);

                    //if (lastKey.Length == 3) lastKey = "00" + lastKey;
                    //else if (lastKey.Length == 4) lastKey = "0" + lastKey;


                    //TabControl_Demande.Items.Clear();
                    //TabControl_Demande = new TabControl();
                    //Demande.Children.Add(TabControl_Demande);

                    // Num_demande.Value = lastKey;

                    //Num_demande.Clear();

                    //Num_demande.Text.Replace(Num_demande.Text, lastKey);


                    Num_demande.Text = ((Int32.Parse(Num_demande.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                    if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                    else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;
                    lastKey = Num_demande.Text;



                    if (connect.State == ConnectionState.Open) connect.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
                }



            }

        }

        private void Ajout_demande_deces_employee_Click(object sender, RoutedEventArgs e)
        {


            Query = "SELECT Matricule FROM Fonctionnaire ";
            Query += " WHERE NomFonct ='" + Nom_fonct_demande_deces_employee.Text + "' AND PrenFonct='" + Prenom_fonct_demande_deces_employee.Text + "'";

            string Matricule = getExcuteScalar(Query);

            Query = "SELECT CodePrime FROM TypePrime ";
            Query += " WHERE DésignationPrime = 'Décès " + Type_deces.Text + "'";

            string CodePrime = getExcuteScalar(Query);

            Query = "SELECT MontantPrime FROM TypePrime";
            Query += " WHERE DésignationPrime='Décès " + Type_deces.Text + "'";

            string Montant = getExcuteScalar(Query);

            Query = "SELECT CompteFonct FROM Fonctionnaire";
            Query += " WHERE Matricule=" + Matricule;

            string CompteFonct = getExcuteScalar(Query);

            string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande_deces_employee.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

            string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment_deces_employee.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



            Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser, NomParent, PrenParent, LienParent, SitFamParent )";
            Query += " VALUES (" + Num_demande_deces_employee.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + CodeUser + ", '" + Nom_demandeur_deces_employee.Text + "', '" + Prenom_demandeur_deces_employee.Text + "' , '" + Lien_parenté.Text + "' , '" + Sit_fam_demandeur_deces_employee.Text + "')";


            try
            {
                if (connect.State == ConnectionState.Closed) connect.Open();
                command = new SqlCommand(Query, connect);
                command.ExecuteNonQuery();
                MessageBox.Show("La demande a été ajouté !");

                Num_demande.Text = ((Int32.Parse(Num_demande.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;
                lastKey = Num_demande.Text;




                Query = "UPDATE Fonctionnaire SET DateDepartDefi ='" + Date_de_event + "' , MotifDepartDefi='Décès' WHERE Matricule=" + Matricule;

                try
                {
                    if (connect.State == ConnectionState.Closed) connect.Open();
                    command = new SqlCommand(Query, connect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Départ définitif pour le fonctionnaire");
                    if (connect.State == ConnectionState.Open) connect.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\nLa base de donnée n'a pas pu être mise à jour !");
                }


                if (connect.State == ConnectionState.Open) connect.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
            }


        }

        private void Ajout_demande_deces_parent_Click(object sender, RoutedEventArgs e)
        {

            Query = "SELECT Matricule FROM Fonctionnaire ";
            Query += " WHERE NomFonct ='" + Nom_fonct_demande_deces.Text + "' AND PrenFonct='" + Prenom_fonct_demande_deces.Text + "'";

            string Matricule = getExcuteScalar(Query);

            Query = "SELECT CodePrime FROM TypePrime ";
            Query += " WHERE DésignationPrime = 'Décès " + Type_deces.Text + "'";

            string CodePrime = getExcuteScalar(Query);

            Query = "SELECT MontantPrime FROM TypePrime";
            Query += " WHERE DésignationPrime='Décès " + Type_deces.Text + "'";

            string Montant = getExcuteScalar(Query);

            Query = "SELECT CompteFonct FROM Fonctionnaire";
            Query += " WHERE Matricule=" + Matricule;

            string CompteFonct = getExcuteScalar(Query);

            string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

            string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



            Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser)";
            Query += " VALUES (" + Num_demande_deces.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + CodeUser + ")";


            try
            {
                if (connect.State == ConnectionState.Closed) connect.Open();
                command = new SqlCommand(Query, connect);
                command.ExecuteNonQuery();
                MessageBox.Show("La demande a été ajouté !");


                Num_demande.Text = ((Int32.Parse(Num_demande.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;
                lastKey = Num_demande.Text;


                if (connect.State == ConnectionState.Open) connect.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
            }



        }

        private void Ajout_demande_Don_Click(object sender, RoutedEventArgs e)
        {

            Query = "SELECT Matricule FROM Fonctionnaire ";
            Query += " WHERE NomFonct ='" + Nom_fonct_demande_Don.Text + "' AND PrenFonct='" + Prenom_fonct_demande_Don.Text + "'";

            string Matricule = getExcuteScalar(Query);

            Query = "SELECT CodePrime FROM TypePrime ";
            Query += " WHERE DésignationPrime = '" + Type_prime_demande.Text + "'";

            string CodePrime = getExcuteScalar(Query);

            Query = "SELECT CompteFonct FROM Fonctionnaire";
            Query += " WHERE Matricule=" + Matricule;

            string CompteFonct = getExcuteScalar(Query);

            string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

            string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



            Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser)";
            Query += " VALUES (" + Num_demande_Don.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant_don.Value + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + CodeUser + ")";


            try
            {
                if (connect.State == ConnectionState.Closed) connect.Open();
                command = new SqlCommand(Query, connect);
                command.ExecuteNonQuery();
                MessageBox.Show("La demande a été ajouté !");

                Num_demande.Text = ((Int32.Parse(Num_demande.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;
                lastKey = Num_demande.Text;


                if (connect.State == ConnectionState.Open) connect.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
            }



        }



        // PARTIE Etat de virement

        private void EtatVir1_Click(object sender, RoutedEventArgs e)
        {

            var application = new Microsoft.Office.Interop.Word.Application();
            var Formulaire = new Microsoft.Office.Interop.Word.Document();

            //string path = "Etat.docx";
            //System.IO.Path.GetFullPath(path);

            Formulaire = application.Documents.Add(Template: @"C:\Users\amine\Documents\Etat.docx");



            foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
            {
                if (field.Code.Text.Contains("Ministere"))
                {
                    try
                    {

                        field.Select();
                        Query = "SELECT Ministere FROM Parametres";
                        application.Selection.TypeText(getExcuteScalar(this.Query));

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }


                }

                else if (field.Code.Text.Contains("Organisme"))
                {
                    try
                    {

                        field.Select();
                        Query = "SELECT Organisme FROM Parametres";
                        application.Selection.TypeText(getExcuteScalar(this.Query));

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }

                else if (field.Code.Text.Contains("CmptSoc"))
                {
                    try
                    {

                        field.Select();
                        Query = "SELECT CompteSocEsi FROM Parametres";
                        application.Selection.TypeText(getExcuteScalar(this.Query));

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }


                }

                else if (field.Code.Text.Contains("CmptTresor"))
                {
                    try
                    {

                        field.Select();
                        Query = "SELECT CompteEsiTresor FROM Parametres";
                        application.Selection.TypeText(getExcuteScalar(this.Query));

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }

                else if (field.Code.Text.Contains("Somme"))
                {
                    try
                    {

                        field.Select();

                        Query = "SELECT SUM(MontantPrime) FROM TypePrime ";
                        Query += " INNER JOIN DemandePrime ON TypePrime.CodePrime = DemandePrime.CodePrime ";
                        Query += " WHERE pv_codepv = 1 AND EtatDem='A'";

                        string somme = getExcuteScalar(this.Query);

                        StringBuilder sb = new StringBuilder();
                        if (somme.Length % 2 == 1) somme = " " + somme;
                        for (int i = 0; i < somme.Length; i++)
                        {
                            if (i % 3 == 0)
                                sb.Append(' ');
                            sb.Append(somme[i]);
                        }

                        application.Selection.TypeText(sb.ToString());

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else if (field.Code.Text.Contains("SomeLettres"))
                {


                    SqlCommand command = new SqlCommand(Query, connect);
                    try
                    {
                        connect.Open();
                        field.Select();
                        Query = "SELECT SUM(MontantPrime) FROM TypePrime ";
                        Query += " INNER JOIN DemandePrime ON TypePrime.CodePrime = DemandePrime.CodePrime ";
                        Query += " WHERE pv_codepv = 1 AND EtatDem='A'";

                        ConvertisseurChiffresLettres convert = new ConvertisseurChiffresLettres();
                        double somme = (double)command.ExecuteScalar();
                        string someLettres = convert.convertion(somme);
                        application.Selection.TypeText(someLettres.ToUpper());
                        connect.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }

                else if (field.Code.Text.Contains("Date"))
                {
                    field.Select();
                    application.Selection.TypeText(date_cheque.Text);
                }

                else if (field.Code.Text.Contains("NumCheque"))
                {
                    field.Select();
                    application.Selection.TypeText(NumCheque.Text);
                }

                else if (field.Code.Text.Contains("Observation"))
                {
                    try
                    {

                        field.Select();
                        Query = "SELECT ObserVir FROM Virement WHERE pv_codepv=1";
                        application.Selection.TypeText(getExcuteScalar(this.Query).ToUpper());

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else if (field.Code.Text.Contains("NumVir"))
                {
                    try
                    {

                        field.Select();
                        Query = "SELECT CodeVir FROM Virement Where pv_codepv =1";

                        string NumVirement = getExcuteScalar(this.Query);
                        if (NumVirement.Length == 1) NumVirement = "00" + NumVirement;
                        else if (NumVirement.Length == 2) NumVirement = "0" + NumVirement;
                        application.Selection.TypeText(NumVirement);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }

            Query = "SELECT CodeVir FROM Virement Where pv_codepv =1";

            string NumVir = getExcuteScalar(this.Query);
            if (NumVir.Length == 1) NumVir = "00" + NumVir;
            else if (NumVir.Length == 2) NumVir = "0" + NumVir;

            Formulaire.SaveAs2(FileName: @".\ProjectTests\EtatVir_" + NumVir + "_" + DateTime.Now.ToString("yy") + ".docx");

            application.Visible = true;
        }


        // PARTIE AVIS & ORDRE


        private void Remplir_Click(object sender, RoutedEventArgs e)
        {
            var application = new Microsoft.Office.Interop.Word.Application();
            var Formulaire = new Microsoft.Office.Interop.Word.Document();

            Formulaire = application.Documents.Add(Template: @"C:\Users\amine\Documents\Avis.docx");



            foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
            {


                if (field.Code.Text.Contains("Name"))
                {

                    field.Select();
                    application.Selection.TypeText(LName.Text);

                }


                else if (field.Code.Text.Contains("PRENOM"))
                {

                    field.Select();
                    application.Selection.TypeText(name.Text);

                }


                else if (field.Code.Text.Contains("CmpTres"))
                {

                    field.Select();
                    Query = "SELECT CompteEsiTresor FROM Parametres";
                    application.Selection.TypeText(getExcuteScalar(this.Query).Substring(0, 7));

                }


                else if (field.Code.Text.Contains("Montant"))
                {

                    field.Select();
                    Query = "SELECT MontantPrime FROM TypePrime ";
                    Query += " INNER JOIN DemandePrime ON TypePrime.CodePrime = DemandePrime.CodePrime ";
                    Query += " WHERE Matricule = 1 AND pv_codepv = 1";

                    string Total = getExcuteScalar(this.Query);

                    // Module pour mettre un espace entre chaque 3 chiffres

                    StringBuilder sb = new StringBuilder();
                    if (Total.Length % 2 == 1) Total = " " + Total;
                    for (int i = 0; i < Total.Length; i++)
                    {
                        if (i % 3 == 0)
                            sb.Append(' ');
                        sb.Append(Total[i]);
                    }

                    application.Selection.TypeText(sb.ToString());

                }


                else if (field.Code.Text.Contains("Clé"))
                {

                    field.Select();
                    Query = "SELECT CompteEsiTresor FROM Parametres";
                    application.Selection.TypeText(getExcuteScalar(this.Query).Substring(8, 2));

                }

                else if (field.Code.Text.Contains("Cle"))
                {

                    field.Select();
                    Query = "SELECT CompteFonct FROM Fonctionnaire WHERE Matricule = 1";
                    application.Selection.TypeText(getExcuteScalar(this.Query).Substring(18, 2));

                }


                else if (field.Code.Text.Contains("CompteFonct"))
                {

                    field.Select();
                    Query = "SELECT CompteFonct FROM Fonctionnaire WHERE NomFonct='" + LName.Text + "' AND PrenFonct='" + name.Text + "'";
                    application.Selection.TypeText(getExcuteScalar(this.Query).Substring(0, 18));


                }


                else if (field.Code.Text.Contains("Motif"))
                {

                    field.Select();
                    Query = "SELECT ObserVir FROM Virement Where pv_codepv =1";
                    application.Selection.TypeText(getExcuteScalar(this.Query));

                }


                else if (field.Code.Text.Contains("Date"))
                {

                    field.Select();
                    application.Selection.TypeText(Date_avis.Text);
                }
            }


            Formulaire.SaveAs2(FileName: @".\ProjectTests\Avis_Ordre_" + LName.Text + "_" + name.Text + ".docx");


            application.Visible = true;

        }


        // PARTIE LISTE


        private void RempListe_Click(object sender, RoutedEventArgs e)
        {
            var application = new Microsoft.Office.Interop.Word.Application();
            var Formulaire = new Microsoft.Office.Interop.Word.Document();




            if (comboBox.SelectedIndex > -1)
            {
                Formulaire = application.Documents.Add(Template: @"C:\Users\chihab\Documents\Liste.docx");

                foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                {

                    if (connect.State == ConnectionState.Closed) { connect.Open(); }

                    int NumPV = codePV(int.Parse(comboBox.Text));

                    Query = "SELECT NomFonct,PrenFonct,MontantPrime,DésignationPrime,CompteFonct FROM DemandePrime ";
                    Query += " INNER JOIN Fonctionnaire ON Fonctionnaire.Matricule=DemandePrime.Matricule";
                    Query += " INNER JOIN TypePrime ON DemandePrime.CodePrime=TypePrime.CodePrime";
                    Query += " WHERE pv_codepv=" + NumPV + " AND EtatDem='A'  ORDER BY NomFonct";

                    command = new SqlCommand(Query, connect);

                    SqlDataReader reader = command.ExecuteReader();


                    if (reader.HasRows)
                    {




                        if (field.Code.Text.Contains("Name"))
                        {
                            field.Select();
                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(0).ToString() + " " + reader.GetValue(1).ToString() + "\n");
                            }
                            reader.Close();
                        }

                        else if (field.Code.Text.Contains("Montant"))
                        {
                            field.Select();

                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(2).ToString() + "\n");
                            }
                            reader.Close();
                        }


                        else if (field.Code.Text.Contains("Motif"))
                        {
                            field.Select();
                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(3).ToString() + "\n");
                            }
                            reader.Close();
                        }


                        else if (field.Code.Text.Contains("Compte"))
                        {
                            field.Select();
                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(4).ToString() + "\n");
                            }
                            reader.Close();
                        }
                        reader.Close();

                        Query = "SELECT CodeVir FROM Virement Where pv_codepv =1";

                        string NumVir = getExcuteScalar(this.Query);
                        if (NumVir.Length == 1) NumVir = "00" + NumVir;
                        else if (NumVir.Length == 2) NumVir = "0" + NumVir;

                        Formulaire.SaveAs2(FileName: @".\ProjectTests\ListeVirement_" + NumVir + "_" + DateTime.Now.ToString("yy") + ".docx");

                    }

                    else if (!reader.HasRows)
                    {
                        MessageBox.Show("Le PV n'existe pas ou est vide !");
                    }

                    if (connect.State == ConnectionState.Open) { connect.Close(); }


                }

                application.Visible = true;
            }
            else
            {
                MessageBox.Show("Selectionner un virement ");
            }
        }



        // PARTIE STATISTIQUES

        private void type_graphe_Initialized(object sender, EventArgs e)
        {
            type_graphe.Items.Add("Histogramme");
            type_graphe.Items.Add("Cercle");
        }

        private void affiche_stats_Click(object sender, RoutedEventArgs e)
        {
            if (type_graphe.Text == "Diagramme")
            {

                Diagram.DEBUT = date_debut_satats;
                Diagram.FIN = date_fin_satats;


                Stats statistiques = new Stats();


                statistiques.WindowState = WindowState.Maximized;
                statistiques.Show();
                statistiques.sourceDiagram.chart.Visibility = Visibility.Visible;

            }

            if (type_graphe.Text == "Cercle")
            {

                Diagram.DEBUT = date_debut_satats;
                Diagram.FIN = date_fin_satats;


                Stats statistiques = new Stats();



                statistiques.WindowState = WindowState.Maximized;
                statistiques.Show();
                statistiques.sourceDiagram.Don.Visibility = Visibility.Visible;

            }

            if (type_graphe.SelectedItem == null)
            {
                MessageBox.Show("Veuillez choisir un type de graphe !");
            }
        }





        // PARTIE EXCEL


        public static int NbTablesExcel;
        public static int NumTableCourante = 1;

        public DataTable dt = new DataTable();

        private void Previous_table_Initialized(object sender, EventArgs e)
        {
            Previous_table.IsEnabled = false;
        }

        private void Next_table_Initialized(object sender, EventArgs e)
        {
            Next_table.IsEnabled = false;
        }

        private void Import_excel_Initialized(object sender, EventArgs e)
        {
            Import_excel.IsEnabled = false;
        }


        public void excelToDataGrid(int NumTableCourante, int NumTable)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            //Static File From Base Path...........
            //Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "TestExcel.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //Dynamic File Using Uploader...........
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            NbTablesExcel = excelBook.Sheets.Count;
            NumTableCourante = NumTable;
            SheetNumber.Text = NumTableCourante.ToString();

            //MessageBox.Show("The Number of sheets is : " + i.ToString());

            try
            {

                //DataTable dt = new DataTable();


                dt.Columns.Clear();
                dt.Rows.Clear();
                dtGrid_Grid.Children.Clear();
                dtGrid = new DataGrid();
                dtGrid.VerticalAlignment = VerticalAlignment.Center;
                dtGrid.HorizontalAlignment = HorizontalAlignment.Center;
                dtGrid_Grid.Children.Add(dtGrid);


                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(NumTableCourante); ;

                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                string strCellData = "";
                double douCellData;
                int rowCnt = 0;
                int colCnt = 0;

                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }


                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch (Exception ex)
                        {

                            douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    dt.Rows.Add(strData.Split('|'));
                }

                dtGrid.ItemsSource = dt.DefaultView;

                excelBook.Close(true, null, null);
                excelApp.Quit();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {


            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".xlsx";
            openfile.Filter = "(.xlsx)|*.xlsx";
            //openfile.ShowDialog();

            var browsefile = openfile.ShowDialog();

            if (browsefile == true)
            {
                txtFilePath.Text = openfile.FileName;


                excelToDataGrid(NbTablesExcel, NumTableCourante);


                Previous_table.IsEnabled = true;
                Next_table.IsEnabled = true;
                Import_excel.IsEnabled = true;

            }
        }

        private void Next_table_Click(object sender, RoutedEventArgs e)
        {




            if (NumTableCourante < NbTablesExcel)
            {

                NumTableCourante++;
                if (0 < NumTableCourante && NumTableCourante <= NbTablesExcel)
                {

                    excelToDataGrid(NbTablesExcel, NumTableCourante);

                }
            }

            else MessageBox.Show("Vous êtes à la dernière table !");


        }

        private void Previous_table_Click(object sender, RoutedEventArgs e)
        {



            if (NumTableCourante > 1)
            {

                NumTableCourante--;
                if (0 < NumTableCourante && NumTableCourante <= NbTablesExcel)
                {
                    excelToDataGrid(NbTablesExcel, NumTableCourante);

                }
            }

            else MessageBox.Show("Vous êtes à la première table !");


        }



        private void Import_excel_Click(object sender, RoutedEventArgs e)
        {
            // C'EST ICI QU'IL FAUT SAVOIR MANIPULER LE DATAGRID

            //foreach ( DataGridRow row in dtGrid.Items )
            //{
            //    string conString = " Data Source = THEPUNISHER; Initial Catalog = OeuvresSociales2; Integrated Security = True";
            //    Query = " INSERT INTO utilisateur ";
            //    Query += " VALUES (@NomUser,@PrenUser,@Login,@MotPasse,@droit)";
            //    using (SqlConnection con = new SqlConnection(conString))
            //    {
            //        using (SqlCommand cmd = new SqlCommand(Query))
            //        {
            //            cmd.Parameters.AddWithValue("@NomUser",row)
            //        }
            //    } 
            //} 



            //DataGridRow row = (DataGridRow)dtGrid.ItemContainerGenerator.ContainerFromIndex(1);



            //BD bd = new BD() ;

            //bd.seConnecter();

            //DataTable dt = new DataTable();

            //dt = ((DataView)dtGrid.ItemsSource).ToTable();


            //DataSet ds = new DataSet();


            //SqlDataAdapter da = new SqlDataAdapter();




            ////SqlCommandBuilder builder = new SqlCommandBuilder(da);

            ////da.UpdateCommand = builder.GetUpdateCommand();

            ////da.Update(ds);




            //da = bd.getDataAdapter("SELECT * FROM utilisateur");

            ////da.Fill(ds, "utilisateur");

            //ds.Tables.Add(dt);

            //SqlCommandBuilder builder = new SqlCommandBuilder(da);

            //da.UpdateCommand = builder.GetUpdateCommand();

            //da.Update(ds);



            ////dtGrid.ItemsSource = ds.DefaultViewManager;


            //path = System.IO.Directory.GetCurrentDirectory();

            //MessageBox.Show(path);

            if (Tables_BDD.Text == "Utilisateur")
            {

                try
                {

                    foreach (DataRow row in dt.Rows)
                    {
                        connect.Open();


                        Query = "INSERT INTO utilisateur VALUES ( ";

                        Query += "'" + row["NomUser"] + "','" + row["PrenUser"] + "','" + row["Login"] + "','" + row["MotPasse"] + "','" + row["droit"] + "'";

                        Query += ")";

                        MessageBox.Show(Query);

                        command = new SqlCommand(Query, connect);

                        command.ExecuteNonQuery();

                        connect.Close();
                    }




                    //Query += row[column].ToString() + ", ";

                    MessageBox.Show("Les données ont été imortés !");


                }



                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



            }


            if (Tables_BDD.Text == "Fonctionnaire")
            {


                try
                {

                    foreach (DataRow row in dt.Rows)
                    {
                        connect.Open();


                        Query = "INSERT INTO Fonctionnaire ( Matricule,NomFonct,PrenFonct,DateRecrut,TelFonct,EmailFonct,CompteFonct,SitFamFonct) ";

                        Query += "VALUES (" + row["Matricule"] + ",'" + row["NomFonct"] + "','" + row["PrenFonct"] + "','" + row["DateRecrut"] + "','";

                        Query += row["TelFonct"] + "','" + row["EmailFonct"] + "'" + row["CompteFonct"] + "','" + row["SitFamFonct"] + "')";

                        MessageBox.Show(Query);

                        command = new SqlCommand(Query, connect);

                        command.ExecuteNonQuery();

                        connect.Close();
                    }




                    //Query += row[column].ToString() + ", ";

                    MessageBox.Show("Les données ont été imortés !");


                }



                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }







        }


        private void Tables_BDD_Initialized(object sender, EventArgs e)
        {
            Tables_BDD.Items.Add("Utilisateur");
            Tables_BDD.Items.Add("Fonctionnaire");
        }





        // PARTIE MISE A JOUR DES INFORMATION DE L'UTILISATEUR





        private void Mon_Compte(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Visible;
            CreerSimpleUser.Visibility = Visibility.Hidden;
            SuppSimpleUser.Visibility = Visibility.Hidden;
            ModifSimpleUser.Visibility = Visibility.Hidden;
        }

        private void Creer_Simple_Utilisateur(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Hidden;
            CreerSimpleUser.Visibility = Visibility.Visible;
            SuppSimpleUser.Visibility = Visibility.Hidden;
            ModifSimpleUser.Visibility = Visibility.Hidden;
        }

        private void Supprimer_Simple_Utilisateur(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Hidden;
            CreerSimpleUser.Visibility = Visibility.Hidden;
            SuppSimpleUser.Visibility = Visibility.Visible;
            ModifSimpleUser.Visibility = Visibility.Hidden;
        }

        private void Modifier_Simple_Utilisateur(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Hidden;
            CreerSimpleUser.Visibility = Visibility.Hidden;
            SuppSimpleUser.Visibility = Visibility.Hidden;
            ModifSimpleUser.Visibility = Visibility.Visible;
        }

        private void Nom_fonct_demande_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Deconnexion(object sender, RoutedEventArgs e)
        {
            LOGIN w = new LOGIN();
            w.Show();
            //this.Close();
            Window window = Window.GetWindow(this);

            window.Close();
        }

        private void Modifier_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void CreerSimpleUtilisateur_Click(object sender, RoutedEventArgs e)
        {
            if (MotPass_b.Password == MotPass_bConf.Password)
            {
                Query = "insert into utilisateur values ";
                Query += "('" + Nom_b.Text + "','" + Prenom_b.Text + "','" + NomUtilisateur_b.Text + "','"+ MotPass_b.Password + "','U')";

                command = new SqlCommand(Query, connect);

                try
                {
                    connect.Open();
                    command.ExecuteNonQuery();
                    MessageBox.Show("L'utilisateur a été créer");
                    connect.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Les mots de passe sont pas identiques, veuillez reconfirmer votre mot de passe.");
            }
 
        }

        private void MotPass_bConf_LostFocus(object sender, RoutedEventArgs e)
        {
            if (MotPass_b.Password == MotPass_bConf.Password)
            {
                MotPass_b.BorderBrush = System.Windows.Media.Brushes.Green;
                MotPass_bConf.BorderBrush = System.Windows.Media.Brushes.Green;
                MessageErreur.Visibility = Visibility.Hidden;
                if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password!= null)&& (NomUtilisateur_b.Text!=null) && (MotPass_bConf.Password != null))
                {
                    CreerSimpleUtilisateur.IsEnabled = true;
                }
            }
            else
            {
                MotPass_b.BorderBrush = System.Windows.Media.Brushes.Red;
                MotPass_bConf.BorderBrush = System.Windows.Media.Brushes.Red;
                MessageErreur.Visibility = Visibility.Visible;
            }
           
        }

        private void Nom_b_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password != null) && (NomUtilisateur_b.Text != null) && (MotPass_bConf.Password != null))
            {
                CreerSimpleUtilisateur.IsEnabled = true;
            }
        }

        private void Prenom_b_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password != null) && (NomUtilisateur_b.Text != null) && (MotPass_bConf.Password != null))
            {
                CreerSimpleUtilisateur.IsEnabled = true;
            }
        }

        private void NomUtilisateur_b_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password != null) && (NomUtilisateur_b.Text != null) && (MotPass_bConf.Password != null))
            {
                CreerSimpleUtilisateur.IsEnabled = true;
            }
        }




        // PARTIE CHIHEB




        private void Ajout_Fonct_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(matricule.Text))
            {
                Fonctionnaire fonc = new Fonctionnaire(int.Parse(matricule.Text), nom.Text, prenom.Text, date.SelectedDate, long.Parse(Ntel.Text), email.Text, compte.Text, code.Text, sitfam.SelectedValue.ToString());
                fonc.Add_fonctionnaire();
            }
            else
            {
                MessageBox.Show("le champ Matricule est vide");
            }
        }

        private void oui_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Ajout_prime_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(prim.Text) && !string.IsNullOrEmpty(montan.Text))
            {
                Prime prime = new Prime(prim.Text, double.Parse(montan.Text));
                prime.Add_Prime();
            }
            else
            {
                MessageBox.Show("Tous les champs sont obligatoire");
            }

        }

        private void Ajout_Banque_Click(object sender, RoutedEventArgs e)
        {
            /*if (!string.IsNullOrEmpty(Des_ban.Text) && !string.IsNullOrEmpty(Adr_ban.Text))
            {
                Banque ban = new Banque(Des_ban.Text, Adr_ban.Text);
                ban.ad;
            }
            else
            {
                MessageBox.Show("Tous les champs sont obligatoire");
            }*/
        }

        private void Upload()
        {
            DataSet ds;
            SqlDataAdapter da;

            BD con = new BD();
            con.seConnecter();
            da = con.getDataAdapter("SELECT        Fonctionnaire.NomFonct AS Nom, Fonctionnaire.PrenFonct AS Prenom, TypePrime.DésignationPrime AS Prime "
                         + " FROM            Fonctionnaire INNER JOIN "
                        + " DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule INNER JOIN "
                         + " TypePrime ON DemandePrime.CodePrime = TypePrime.CodePrime AND DemandePrime.pv_codepv IS NULL");

            ds = new DataSet();
            da.Fill(ds, "tabel1");
            dataGrid.ItemsSource = ds.Tables["tabel1"].DefaultView;
            con.seConnecter();


        }

        private void list_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Would you like to go to hell ", "Attention!!", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:
                    {
                        BD con = new BD();
                        con.seConnecter();
                        string query = "UPDATE       DemandePrime  SET   pv_codepv =1000   WHERE    pv_codepv IS NULL ";
                        con.executerRequete(query);
                        Upload();
                        Upload_traitement();
                        con.seConnecter();

                    }
                    break;
                case MessageBoxResult.No:
                    Upload();
                    break;
            }
        }

        private void Upload_traitement()
        {
            DataSet ds1;
            SqlDataAdapter da1;
            BD con = new BD();
            con.seConnecter();
            da1 = con.getDataAdapter("SELECT        DemandePrime.NumDem AS Demande, Fonctionnaire.NomFonct AS NOM , Fonctionnaire.PrenFonct AS PRENOM , TypePrime.DésignationPrime AS Prime "
                         + " FROM            Fonctionnaire INNER JOIN "
                         + " DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule INNER JOIN"
                         + " TypePrime ON DemandePrime.CodePrime = TypePrime.CodePrime AND DemandePrime.pv_codepv = 1000");

            ds1 = new DataSet();
            da1.Fill(ds1, "tabel");
            dataGrid1.ItemsSource = ds1.Tables["tabel"].DefaultView;
            con.seConnecter();
        }



        private void decision_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BD con = new BD();
            con.seConnecter();
            ComboBox combo = sender as ComboBox;
            DataRowView dataRow = (DataRowView)dataGrid1.SelectedItem;
            int numdem = int.Parse(dataRow.Row.ItemArray[0].ToString());
            //int index = dataGrid1.CurrentCell.Column.DisplayIndex;

            char cellValue = char.Parse(combo.SelectedValue.ToString());
            string query = "UPDATE       DemandePrime SET                EtatDem = '" + cellValue + "'     WHERE        (NumDem = " + numdem + ") ";
            con.executerRequete(query);
            con.seConnecter();
        }



        private void dataGrid1_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var editedTextbox = e.EditingElement as TextBox;

            if (editedTextbox != null)
            {
                BD con = new BD();
                con.seConnecter();
                DataRowView dataRow = (DataRowView)dataGrid1.SelectedItem;
                int numdem = int.Parse(dataRow.Row.ItemArray[0].ToString());
                string query = "UPDATE       DemandePrime SET                MotifEtat  = '" + editedTextbox.Text + "'     WHERE        (NumDem = " + numdem + ") ";
                con.executerRequete(query);
                con.seConnecter();
            }

        }

        private void PV_Click(object sender, RoutedEventArgs e)
        {
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT  TOP (1) CodePV FROM PV  ORDER BY CodePV DESC");
            if (dr.Read())
            {
                int a = DateTime.Today.Year % 100 * 1000;
                int code = Convert.ToInt32(dr[0]);

                if (code >= a)
                {
                    code++;
                }
                else
                {
                    code = a;
                }
                MessageBoxResult result = MessageBox.Show("Une fois que vous créez ce PV vous n aurez pas le droit de changer les décisions ", "Attention!!", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        {
                            PV p = new PV(code);
                            p.Add_PV();
                            con.seConnecter();
                            string query = "UPDATE       DemandePrime  SET   pv_codepv = " + code + "   WHERE    pv_codepv = 1000 AND EtatDem != 'I' ";
                            con.executerRequete(query);
                            Upload_traitement();

                        }
                        break;
                    case MessageBoxResult.No:
                        Upload();

                        break;
                }

            }





            con.seDeconnecter();

        }



        private void button1_Click(object sender, RoutedEventArgs e)
        {
            BD con = new BD();
            con.seConnecter();
            int code = 0;
            SqlDataReader dr = con.getResultatRequete("SELECT  TOP (1) CodeVir FROM Virement  ORDER BY CodeVir DESC");
            if (dr.Read())
            {
                int a = DateTime.Today.Year % 100 * 1000;
                code = Convert.ToInt32(dr[0]);

                if (code >= a)
                {
                    code++;
                }
                else
                {
                    code = a;
                }
                con.seDeconnecter();
            }
            MessageBoxResult result = MessageBox.Show("Voulez vous vraiment creez ce virement ? ", "Attention!!", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:
                    {
                        con.seConnecter();

                        SqlDataReader dr1 = con.getResultatRequete("SELECT * FROM Parametres");

                        if (dr1.Read() && N_PV.SelectedIndex > -1)
                        {
                            BD con1 = new BD();
                            con1.seConnecter();
                            MessageBox.Show(code.ToString());
                            string query = " INSERT INTO Virement "
                             + " (CodeVir, pv_codepv, DateCreatVir, CodeUser, MinistereVir, OrganismeVir, CompteSocVir, CompteEsiVir, BenefVir, ObserVir) "
                            + " VALUES(" + code + "," + int.Parse(N_PV.Text) + ",'" + DateTime.Today + "',01,'" + dr1[0].ToString() + "','" + dr1[1].ToString() + "','" + dr1[5].ToString() + "','" + dr1[6].ToString() + "','" + dr1[7].ToString() + "','" + observation.Text + "')";

                            observation.Text = query;
                            con1.seDeconnecter();

                        }
                    }
                    break;
                case MessageBoxResult.No:
                    Upload();

                    break;
            }
        }

        private void N_PV_DropDownOpened(object sender, EventArgs e)
        {
            N_PV.Items.Clear();
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT   CodePV FROM PV WHERE Virement IS NULL  ORDER BY CodePV DESC");
            while (dr.Read())
            {
                N_PV.Items.Add((int)dr[0]);
            }
            con.seDeconnecter();
        }

        private void comboBox_DropDownOpened(object sender, EventArgs e)
        {
            combo.Items.Clear();
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT        CodeVir   FROM            Virement     ORDER BY CodeVir DESC");
            while (dr.Read())
            {
                combo.Items.Add((int)dr[0]);
            }
            con.seDeconnecter();
        }



        private int codePV(int codevir)
        {
            BD con = new BD();
            con.seConnecter();
            int code = 0;
            SqlDataReader dr = con.getResultatRequete("SELECT        pv_codepv  FROM            Virement   WHERE        CodeVir = " + codevir);
            if (dr.Read())
            {
                code = (int)dr[0];
            }
            con.seDeconnecter();
            return code;


        }

        private void comboBox_DropDownOpened_1(object sender, EventArgs e)
        {
            comboBox.Items.Clear();
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT        CodeVir   FROM            Virement     ORDER BY CodeVir DESC");
            while (dr.Read())
            {
                comboBox.Items.Add((int)dr[0]);
            }
            con.seDeconnecter();
        }








    }
}
