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
using System.Net.Mail;
using System.Configuration;
using System.Net;
using NewEmployeeDLL;

namespace EmailSetup
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        EmployeeClass TheEmployeeClass = new EmployeeClass();

        FindActiveEmployeesDataSet TheFindActiveEmployeesDataSet = new FindActiveEmployeesDataSet();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSendEmail_Click(object sender, RoutedEventArgs e)
        {
            string strMailTo = "tholmes@bluejaycommunications.com";
            string strSubject = "Active Employee Roster --- Do Not Reply";
            string strMessage = "<h1>Active Employee Roster</h1>";
            int intCounter;
            int intNumberOfRecords;

            try
            {
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Employee ID</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Phone Number</b></td>";
                strMessage += "<td><b>Home Office</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                TheFindActiveEmployeesDataSet = TheEmployeeClass.FindActiveEmployees();

                intNumberOfRecords = TheFindActiveEmployeesDataSet.FindActiveEmployees.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strMessage += "<tr>";
                    strMessage += "<td>" + Convert.ToString(TheFindActiveEmployeesDataSet.FindActiveEmployees[intCounter].EmployeeID) + "</td>";
                    strMessage += "<td>" + TheFindActiveEmployeesDataSet.FindActiveEmployees[intCounter].FirstName + "</td>";
                    strMessage += "<td>" + TheFindActiveEmployeesDataSet.FindActiveEmployees[intCounter].LastName + "</td>";
                    strMessage += "<td>" + TheFindActiveEmployeesDataSet.FindActiveEmployees[intCounter].PhoneNumber + "</td>";
                    strMessage += "<td>" + TheFindActiveEmployeesDataSet.FindActiveEmployees[intCounter].HomeOffice + "</td>";
                    strMessage += "</tr>";
                }
                strMessage += "</table>";

                bool blnEmailSet = SendEmail(strMailTo, strSubject, strMessage);

                if (blnEmailSet == false)
                    throw new Exception();

                MessageBox.Show("Sent");
            }
            catch (Exception)
            {
                MessageBox.Show("Error");
            }
        }
        public bool SendEmail(string mailTo, string subject, string message)
        {
            try
            {

                MailMessage mailMessage = new MailMessage("toolreport@bluejaycommunications.com", mailTo, subject, message);
                mailMessage.IsBodyHtml = true;
                mailMessage.BodyEncoding = Encoding.UTF8;
                mailMessage.SubjectEncoding = Encoding.UTF8;

                SmtpClient smtpClient = new SmtpClient("192.168.0.240", 25);
                //smtpClient.Credentials = CredentialCache.DefaultNetworkCredentials;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.EnableSsl = false;
                smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                //smtpClient.Credentials = new NetworkCredential("bluejaycommunicationsit@gmail.com", "Bluejay2017!");
                smtpClient.Send(mailMessage);
                return true;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString());

                return false;
            }


        }
    }
}
