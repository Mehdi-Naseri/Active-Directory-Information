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
using System.IO;

using System.DirectoryServices.AccountManagement;
using Excel = Microsoft.Office.Interop.Excel;



namespace ActiveDirectoryInformation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private enum EunmDataGrid { empty, users, computers };
        private EunmDataGrid EnumDataGrid1 = EunmDataGrid.empty;
        private static string stringDomainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;

        private void buttonShow_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (stringDomainName != null)
                {
                    datagridResult.ItemsSource = null;
                    if (radiobuttonUsers.IsChecked == true)
                    {
                        ShowUsers();

                    }
                    else if (radiobuttonComputers.IsChecked == true)
                    {
                        ShowComputers();
                    }
                }
                else
                {
                    MessageBox.Show("Your computer is not a member of domain", "Active Directory Users", MessageBoxButton.OK, MessageBoxImage.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void ShowComputers()
        {
            Computers1.Clear();
            int intCounter = 0;
            PrincipalContext PrincipalContext1 = new PrincipalContext(ContextType.Domain, stringDomainName);
            ComputerPrincipal ComputerPrincipal1 = new ComputerPrincipal(PrincipalContext1);
            PrincipalSearcher search = new PrincipalSearcher(ComputerPrincipal1);
            foreach (ComputerPrincipal result in search.FindAll())
            {
                Computer Computer1 = new Computer(result.SamAccountName, result.DisplayName, result.Name, result.Description, result.Enabled, result.LastLogon);
                Computers1.Add(Computer1);
                intCounter++;
            }
            search.Dispose();
            datagridResult.ItemsSource = Computers1;
            datagridResult.Items.Refresh();
            MessageBox.Show(intCounter + " computers. ");
            EnumDataGrid1 = EunmDataGrid.computers;
        }

        private void ShowUsers()
        {
            //Check password
            Boolean boolPass;
            //Check groups
            Boolean boolGroup;
            Users1.Clear();
            int intCounter = 0;
            PrincipalContext PrincipalContext1 = new PrincipalContext(ContextType.Domain, stringDomainName);
            UserPrincipal UserPrincipal1 = new UserPrincipal(PrincipalContext1);
            PrincipalSearcher search = new PrincipalSearcher(UserPrincipal1);

            foreach (UserPrincipal result in search.FindAll())
            {
                //Check criteria
                boolPass = false;
                boolGroup = false;
                //Check default pass
                if (checkBoxPass.IsChecked == true)
                {
                    if (PrincipalContext1.ValidateCredentials(result.SamAccountName, PasswordBoxPass.Password))
                    {
                        boolPass = true;
                    }
                    else
                    {
                        boolPass = false;
                    }
                }
                else
                {
                    boolPass = true;
                }
                //Check group
                if (comboBoxGroups.SelectedIndex >= 0)
                {
                    PrincipalSearchResult<Principal> PrincipalSearchResults1 = result.GetGroups();
                    foreach (Principal PrincipalSearchResult1 in PrincipalSearchResults1)
                    {
                        if (PrincipalSearchResult1.Name == comboBoxGroups.SelectedValue.ToString())
                        {
                            boolGroup = true;
                            break;
                        }
                    }
                }
                else
                {
                    boolGroup = true;
                }
                //Add user
                if (boolPass && boolGroup)
                {
                    User User1 = new User(result.SamAccountName, result.DisplayName, result.Name, result.GivenName, result.Surname,
                       result.Description, result.Enabled, result.LastLogon);
                    Users1.Add(User1);
                    intCounter++;
                }
            }
            search.Dispose();
            datagridResult.ItemsSource = Users1;
            datagridResult.Items.Refresh();
            MessageBox.Show(intCounter + " users. ");
            EnumDataGrid1 = EunmDataGrid.users;
        }

        private void buttonExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (EnumDataGrid1 != EunmDataGrid.empty)
                {
                    if (radiobuttonExcel.IsChecked == true)
                    {
                        Microsoft.Win32.SaveFileDialog SaveFileDialog1 = new Microsoft.Win32.SaveFileDialog();
                        SaveFileDialog1.Filter = "Excel Workbook (*.xls)|*.xls";
                        if ((bool)SaveFileDialog1.ShowDialog())
                        {
                            if (EnumDataGrid1 == EunmDataGrid.users)
                            {
                                ExportUserstoExcel(SaveFileDialog1.FileName);
                            }
                            else if (EnumDataGrid1 == EunmDataGrid.computers)
                            {
                                ExportComputerstoExcel(SaveFileDialog1.FileName);
                            }
                        }

                    }
                    else
                    {
                        Microsoft.Win32.SaveFileDialog SaveFileDialog1 = new Microsoft.Win32.SaveFileDialog();
                        SaveFileDialog1.Filter = "Comma-Seprated Value (*.csv)|*.csv";
                        if ((bool)SaveFileDialog1.ShowDialog())
                        {
                            if (EnumDataGrid1 == EunmDataGrid.users)
                            {
                                ExportUserstoCSV(SaveFileDialog1.FileName);
                            }
                            else if (EnumDataGrid1 == EunmDataGrid.computers)
                            {
                                ExportComputerstoCSV(SaveFileDialog1.FileName);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("First click on show button", "Active Directory Users", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void ExportUserstoExcel(string stringFileName)
        {
            Excel._Application ExcelApplication;
            Excel.Workbook ExcelWorkbook;
            Excel.Worksheet ExcelWorksheet;
            object objectMisValue = System.Reflection.Missing.Value;
            Excel.Range ExcelRangeCellinstance;
            ExcelApplication = new Excel.Application();
            ExcelWorkbook = ExcelApplication.Workbooks.Add(objectMisValue);

            ExcelWorksheet = (Excel.Worksheet)ExcelWorkbook.Worksheets.get_Item(1);
            ExcelApplication.DisplayAlerts = false;
            ExcelRangeCellinstance = ExcelWorksheet.get_Range("A1", Type.Missing);
            int intRow = 1;
            int intColumn = 1;
            foreach (string string1 in User.StringArrayUesrProperties)
            {
                ExcelWorksheet.Cells[intRow, intColumn] = string1;
                intColumn++;
            }
            intRow++;
            foreach (User User1 in Users1)
            {
                intColumn = 1;
                foreach (string string1 in User1.Properties())
                {
                    ExcelWorksheet.Cells[intRow, intColumn] = string1;
                    intColumn++;
                }
                intRow++;
            }
            //Highlight first row
            Excel.Range ExcelRange1 = ExcelWorksheet.get_Range("A1", Type.Missing);
            ExcelRange1.EntireRow.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            ExcelRange1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
            ExcelRange1.EntireRow.Font.Size = 14;
            ExcelRange1.EntireRow.AutoFit();
            //Save Excel
            ExcelWorkbook.SaveAs(stringFileName, Excel.XlFileFormat.xlWorkbookNormal, objectMisValue, objectMisValue, objectMisValue, objectMisValue, Excel.XlSaveAsAccessMode.xlExclusive, objectMisValue, objectMisValue, objectMisValue, objectMisValue, objectMisValue);
            ExcelWorkbook.Close();
            MessageBox.Show("Saved Successfully", "Active Directory", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExportUserstoCSV(string stringFileName)
        {
            //   File FileStream1 = new System.IO.File();
            StringBuilder StringBuilder1 = new StringBuilder(null);
            foreach (string string1 in User.StringArrayUesrProperties)
            {
                if (StringBuilder1.Length == 0)
                    StringBuilder1.Append(string1);
                StringBuilder1.Append(',' + string1);
            }
            StringBuilder1.AppendLine();
            foreach (User User1 in Users1)
            {
                StringBuilder StringBuilderTemp = new StringBuilder(null);
                foreach (string string1 in User1.Properties())
                {
                    if (StringBuilderTemp.Length == 0)
                        StringBuilderTemp.Append(string1);
                    StringBuilderTemp.Append(',' + string1);
                }
                //   StringBuilder1.AppendLine();
                StringBuilder1.AppendLine(StringBuilderTemp.ToString());
            }
            File.WriteAllText(stringFileName, StringBuilder1.ToString(), Encoding.UTF8);
            MessageBox.Show("Saved Successfully", "Active Directory", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExportComputerstoExcel(string stringFileName)
        {
            Excel._Application ExcelApplication;
            Excel.Workbook ExcelWorkbook;
            Excel.Worksheet ExcelWorksheet;
            object objectMisValue = System.Reflection.Missing.Value;
            Excel.Range ExcelRangeCellinstance;
            ExcelApplication = new Excel.Application();
            ExcelWorkbook = ExcelApplication.Workbooks.Add(objectMisValue);

            ExcelWorksheet = (Excel.Worksheet)ExcelWorkbook.Worksheets.get_Item(1);
            ExcelApplication.DisplayAlerts = false;
            ExcelRangeCellinstance = ExcelWorksheet.get_Range("A1", Type.Missing);
            int intRow = 1;
            int intColumn = 1;
            foreach (string string1 in Computer.StringArrayComputerProperties)
            {
                ExcelWorksheet.Cells[intRow, intColumn] = string1;
                intColumn++;
            }
            intRow++;
            foreach (Computer Computer1 in Computers1)
            {
                intColumn = 1;
                foreach (string string1 in Computer1.Properties())
                {
                    ExcelWorksheet.Cells[intRow, intColumn] = string1;
                    intColumn++;
                }
                intRow++;
            }
            //Highlight first row
            Excel.Range ExcelRange1 = ExcelWorksheet.get_Range("A1", Type.Missing);
            ExcelRange1.EntireRow.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            ExcelRange1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
            ExcelRange1.EntireRow.Font.Size = 14;
            ExcelRange1.EntireRow.AutoFit();
            //Save Excel
            ExcelWorkbook.SaveAs(stringFileName, Excel.XlFileFormat.xlWorkbookNormal, objectMisValue, objectMisValue, objectMisValue, objectMisValue, Excel.XlSaveAsAccessMode.xlExclusive, objectMisValue, objectMisValue, objectMisValue, objectMisValue, objectMisValue);
            ExcelWorkbook.Close();
            MessageBox.Show("Saved Successfully", "Active Directory", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExportComputerstoCSV(string stringFileName)
        {
            //   File FileStream1 = new System.IO.File();
            StringBuilder StringBuilder1 = new StringBuilder(null);
            foreach (string string1 in Computer.StringArrayComputerProperties)
            {
                if (StringBuilder1.Length == 0)
                    StringBuilder1.Append(string1);
                StringBuilder1.Append(',' + string1);
            }
            StringBuilder1.AppendLine();
            foreach (Computer Computer1 in Computers1)
            {
                StringBuilder StringBuilderTemp = new StringBuilder(null);
                foreach (string string1 in Computer1.Properties())
                {
                    if (StringBuilderTemp.Length == 0)
                        StringBuilderTemp.Append(string1);
                    StringBuilderTemp.Append(',' + string1);
                }
                //   StringBuilder1.AppendLine();
                StringBuilder1.AppendLine(StringBuilderTemp.ToString());
            }
            File.WriteAllText(stringFileName, StringBuilder1.ToString(), Encoding.UTF8);
            MessageBox.Show("Saved Successfully", "Active Directory", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #region Users
        // public List<string> listStringUesrPropertie = new List<string> { "SamAccountName", "DisplayName", "Name", "GivenName", "Surname", "Description", "Enabled", "LastLogon" };
        public Users Users1 = new Users();
        public class Users : List<User> { }
        public class User
        {
            public String SamAccountName { get; set; }
            public String DisplayName { get; set; }
            public String Name { get; set; }
            public String GivenName { get; set; }
            public String Surname { get; set; }
            public String Description { get; set; }
            public Boolean? Enabled { get; set; }
            public DateTime? LastLogon { get; set; }

            public User(String SamAccountName, String DisplayName, String Name, String GivenName, String Surname, String Description,
                Boolean? Enabled, DateTime? LastLogon)
            {
                this.SamAccountName = SamAccountName;
                this.DisplayName = DisplayName;
                this.Name = Name;
                this.GivenName = GivenName;
                this.Surname = Surname;
                this.Description = Description;
                this.Enabled = Enabled;
                this.LastLogon = LastLogon;
            }
            public List<string> Properties()
            {
                return new List<string> { SamAccountName, DisplayName, Name, GivenName, Surname, Description, Enabled.ToString(), LastLogon.ToString() };
            }
            public int UserPropertiesTotal = 8;
            public static string[] StringArrayUesrProperties = { "SamAccountName", "DisplayName", "Name", "GivenName", "Surname", "Description", "Enabled", "LastLogon" };
        }
        #endregion

        #region Computers
        //public List<string> listStringUesrPropertie = new List<string> { "SamAccountName", "DisplayName", "Name", "GivenName", "Surname", "Description", "Enabled", "LastLogon" };
        public Computers Computers1 = new Computers();
        public class Computers : List<Computer> { }
        public class Computer
        {
            public String SamAccountName { get; set; }
            public String DisplayName { get; set; }
            public String Name { get; set; }
            public String Description { get; set; }
            public Boolean? Enabled { get; set; }
            public DateTime? LastLogon { get; set; }

            public Computer(String SamAccountName, String DisplayName, String Name, String Description,
                Boolean? Enabled, DateTime? LastLogon)
            {
                this.SamAccountName = SamAccountName;
                this.DisplayName = DisplayName;
                this.Name = Name;
                this.Description = Description;
                this.Enabled = Enabled;
                this.LastLogon = LastLogon;
            }
            public List<string> Properties()
            {
                return new List<string> { SamAccountName, DisplayName, Name, Description, Enabled.ToString(), LastLogon.ToString() };
            }
            public int UserPropertiesTotal = 6;
            public static string[] StringArrayComputerProperties = { "SamAccountName", "DisplayName", "Name", "Description", "Enabled", "LastLogon" };
        }
        #endregion

        private void windowMain_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                string stringDomainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
                if (stringDomainName != null)
                {
                    PrincipalContext PrincipalContext1 = new PrincipalContext(ContextType.Domain, stringDomainName);
                    GroupPrincipal GroupPrincipal1 = new GroupPrincipal(PrincipalContext1);
                    PrincipalSearcher search = new PrincipalSearcher(GroupPrincipal1);

                    foreach (GroupPrincipal GroupPrincipal2 in search.FindAll())
                    {
                        comboBoxGroups.Items.Add(GroupPrincipal2.Name);
                    }
                    comboBoxGroups.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("Content", System.ComponentModel.ListSortDirection.Descending));
                }
                else
                {
                    MessageBox.Show("Your computer is not a member of domain", "Active Directory Users", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); 
            }
        }

        private void radiobuttonComputers_Checked(object sender, RoutedEventArgs e)
        {
            comboBoxGroups.IsEnabled = false;
            checkBoxPass.IsEnabled = false;
            PasswordBoxPass.IsEnabled = false;
        }

        private void radiobuttonComputers_Unchecked(object sender, RoutedEventArgs e)
        {
            comboBoxGroups.IsEnabled = true;
            checkBoxPass.IsEnabled = true;
         //   PasswordBoxPass.IsEnabled = true;
        }

        private void checkBoxPass_Checked(object sender, RoutedEventArgs e)
        {
            PasswordBoxPass.IsEnabled = true;
        }

        private void checkBoxPass_Unchecked(object sender, RoutedEventArgs e)
        {
            PasswordBoxPass.IsEnabled = false;
        }
    }
}
