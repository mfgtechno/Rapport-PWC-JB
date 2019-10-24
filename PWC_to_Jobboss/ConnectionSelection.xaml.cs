using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
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

namespace PWC_to_Jobboss
{
    /// <summary>
    /// Interaction logic for ConnectionSelection.xaml
    /// </summary>
    public partial class ConnectionSelection : Window
    {
        public string SelectedDatabase { get; set; }
        public string SelectedInstance { get; set; }

        public ConnectionSelection()
        {
            InitializeComponent();

            List<string> databaseNames = new List<string>();
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\JobBOSS\Common\InstallRoot"))
                //using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\JavaSoft\Auto Update"))
                {
                    if (key != null)
                    {
                        Object o = key.GetValue("JobBOSSServerPath");
                        if (o != null)
                        {
                            string[] lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(o as string, "MachineInformation.ini"));

                            foreach (string s in lines)
                            {
                                if (!string.IsNullOrWhiteSpace(s) && s.ToLower().StartsWith("machinename="))
                                {
                                    string instance = s.ToLower().Replace("machinename=", "");

                                    using (SqlConnection cn = new SqlConnection(string.Concat("Data Source=", instance, ";Initial Catalog=master;user id=support;password=lonestar;MultipleActiveResultSets=True;")))
                                    {
                                        cn.Open();
                                        using (SqlCommand cmd = new SqlCommand("select name from sys.databases where name not in ('master','tempdb','model','msdb')", cn))
                                        {
                                            SqlDataReader rs = cmd.ExecuteReader();

                                            while (rs.Read())
                                                databaseNames.Add(rs[0].ToString());
                                        }
                                        cn.Close();
                                    }

                                    SelectedInstance = instance;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Une erreur est survenue lors du chargement des connexions. Aucune valeur n'a été trouvée.", "Attention", MessageBoxButton.OK, MessageBoxImage.Error);
                            this.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Une erreur est survenue lors du chargement des connexions. La clé de registre est absente.", "Attention", MessageBoxButton.OK, MessageBoxImage.Error);
                        this.Close();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Une erreur est survenue lors du chargement des connexions", "Attention", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
            }

            if (databaseNames.Count == 0)
            {
                MessageBox.Show("Aucune connexion n'a été trouvée malgré la présence de la clé de registre.", "Attention", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
            }

            foreach (string databaseName in databaseNames)
            {
                Label label = new Label
                {
                    Content = databaseName,
                    BorderBrush = Brushes.Black,
                    BorderThickness = new Thickness(2),
                    Margin = new Thickness(10, 10, 10, 0)
                };

                label.MouseDoubleClick += Label_MouseDoubleClick;
                MainPanel.Children.Add(label);
            }
        }

        private void Label_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Label label = sender as Label;
            if (label != null && label.Content != null && !string.IsNullOrWhiteSpace(label.Content.ToString()))
            {
                this.SelectedDatabase = label.Content.ToString();
                this.DialogResult = true;
                this.Close();
            }
        }
    }
}
