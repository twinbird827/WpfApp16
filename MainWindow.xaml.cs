using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
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

namespace WpfApp16
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string IPADDRESS = "";

        private const string DC1 = "";

        private const string DC2 = "";

        private const string USERNAME = "";

        private const string PASSWORD = "";

        public MainWindow()
        {
            InitializeComponent();

            DataContext = Items;
        }

        public ObservableCollection<string> Items = new ObservableCollection<string>();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var root = new DirectoryEntry($"LDAP://{IPADDRESS}/ DC={DC1},DC={DC2}", USERNAME, PASSWORD))
                using (var se = new DirectorySearcher(root, "(objectCategory=Person)"))
                {
                    Items.Clear();
                    foreach (SearchResult res in se.FindAll())
                    {
                        var de = res.GetDirectoryEntry();
                        Items.Add($"{GetProperties(de.Properties)}");
                    }
                    Console.WriteLine($"End.");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string GetProperties(PropertyCollection pc)
        {
            var names = new[]
            {
                "accountExpires",
                "adminCount",
                "badPasswordTime",
                "badPwdCount",
                "cn",
                "codePage",
                "comment",
                "countryCode",
                "dSCorePropagationData",
                "description",
                "displayName",
                "distinguishedName",
                "givenName",
                "initials",
                "instanceType",
                "isCriticalSystemObject",
                "l",
                "lastLogoff",
                "lastLogon",
                "lastLogonTimestamp",
                "logonCount",
                "logonHours",
                "mSMQDigests",
                "mSMQSignCertificates",
                "mail",
                "managedObjects",
                "memberOf",
                "msDS-SupportedEncryptionTypes",
                "msNPAllowDialin",
                "nTSecurityDescriptor",
                "name",
                "objectCategory",
                "objectClass",
                "objectGUID",
                "objectSid",
                "postalCode",
                "primaryGroupID",
                "pwdLastSet",
                "sAMAccountName",
                "sAMAccountType",
                "servicePrincipalName",
                "sn",
                "st",
                "telephoneNumber",
                "uSNChanged",
                "uSNCreated",
                "userAccountControl",
                "userParameters",
                "userPrincipalName",
                "whenChanged",
                "whenCreated",
            };

            var sb = new StringBuilder();
            try
            {
                Console.WriteLine($"[{DateTime.Now.ToString("mm:ss.fffff")}] Started {pc["displayName"].Value}");
                foreach (var name in names)
                {
                    //sb.Append($"{name}:");
                    sb.Append(pc[name].Value);
                    sb.Append($"\t");
                }
//                foreach (var name in pc.PropertyNames)
//                {
//                    sb.Append($"{name}\t");
////                    sb.Append(pc[name.ToString()].Value);
//                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return sb.ToString();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            using (var context = new PrincipalContext(ContextType.Domain, $"{DC1}.{DC2}"))
            {
                if (context.ValidateCredentials(USER.Text, PASS.Text))
                    MessageBox.Show("OK");
                else
                    MessageBox.Show("NG");
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var sb = new StringBuilder();
            foreach (var s in Items)
            {
                sb.AppendLine(s);
            }
            Clipboard.SetText(sb.ToString());
        }
    }
}
