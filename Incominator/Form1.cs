using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net;
using System.IO;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using MimeKit.Utils;

namespace Incominator
{
    public partial class Incominator : Form
    {
        public Incominator()
        {
            InitializeComponent();
            UpdateInformation();

        }
        const double Wage = 14.50;
        const double Loan = 218.39;
        const double gasPerMonth = 50;//36.25;
        const double Expenses = Loan + gasPerMonth;

        double TotalHours = 0;
        double TotalEarnings = 0;
        double TotalNetPay = 0;
        double AverageHours = 0;
        double AverageEarnings = 0;
        double AverageNetPay = 0;

        private void btnEnter_Click(object sender, EventArgs e)
        {
            try
            {
                // File name with the database inside
                string dbFileName = @"C:\Users\cjbra\Desktop\Incominator\Database\Income.accdb";
                string dbConnectionInfo = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + dbFileName + "; Persist Security Info=False";
                //An OleDbConnection object represents a unique connection to a data source
                OleDbConnection dbConnection = new OleDbConnection();
                dbConnection.ConnectionString = dbConnectionInfo;
                dbConnection.Open();

                // Not sure if I want to use DateTime information for anything yet.
                // DateTime TodaysDate = DateTime.Now;

                // Variables that get the entered information to be stored in the database
                double Hours = double.Parse(txtHours.Text);
                double Earnings = double.Parse(txtEarnings.Text);
                double NetPay = double.Parse(txtNetPay.Text);
                double Tax = Earnings - NetPay;

                // Inserts the Hour, Earnings, NetPay, and Tax values into the database.
                string insertString = "Insert into IncomeActual ([Hours], [Earnings], [NetPay], [Tax]) values(" + 
                    Hours + ", " + Earnings + ", " + NetPay + ", " + Tax + ")";

                OleDbCommand sqlCommand;
                sqlCommand = new OleDbCommand();
                //give the sqlCommand the connection.
                sqlCommand.Connection = dbConnection;
                //give the sqlCommand the sql statement
                sqlCommand.CommandText = insertString;
                //  execute sql command
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("The Data is Stored!");
                // Call Close.Remember this is always a good practice.
                dbConnection.Close();

                // Updates the form information when a new entry is added
                UpdateInformation();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in Enter Button: \n" + ex);

            }
        }
        private void UpdateInformation()
        {
            try
            {
                /* Resets all variables on the update start to stop the variables
                 * from adding twice and scewing the results.
                 */
                resetVariables();
                string dbFileName = @"C:\Users\cjbra\Desktop\Incominator\Database\Income.accdb";
                string dbConnectionInfo = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + dbFileName + "; Persist Security Info=False";
                //An OleDbConnection object represents a unique connection to a data source
                OleDbConnection dbConnection = new OleDbConnection();

                dbConnection.ConnectionString = dbConnectionInfo;
                // Open the connection in a try/catch block. 
                dbConnection.Open();
                //Step:1 construct our "select" statement, which will be used to read the data from the database
                string selectString = "Select * from IncomeActual";
                //create  a sql command to execute over our database
                OleDbCommand sqlCommand = new OleDbCommand();
                //give the sqlCommand the connection.
                sqlCommand.Connection = dbConnection;
                //give the sqlCommand the sql statement
                sqlCommand.CommandText = selectString;
                // Create and execute the DataReader, writing the result in the textboxs
                OleDbDataReader reader = sqlCommand.ExecuteReader();

                
                int Counter = 0;
                double intEntryMonths = 0;
                double EarningsRemainder = 0;
                double NetPayRemainder = 0;
                double HoursRemainder = 0;
                // While there is information to be read from the database
                while(reader.HasRows)
                {
                    // While the reader is reading information from the database
                    while (reader.Read())
                    {
                        // Adds the database information for each entry to the total variables
                        TotalHours += double.Parse(reader.GetValue(1).ToString());
                        TotalEarnings += double.Parse(reader.GetValue(2).ToString());
                        TotalNetPay += double.Parse(reader.GetValue(3).ToString());
                        // Adds 1 to the counter to be used in the if statement
                        Counter++;
                        // If the counter variable is equal to 4
                        if(Counter >= 4)
                        {
                            // Adds 1 to the amount of entries per month
                            intEntryMonths++;
                            /* Divides the total entry for 4 data entries to get the average
                             * per month. resets the totals and counter for the next itteration.
                             */
                            AverageEarnings += Math.Round(TotalEarnings / 4, 2);
                            AverageNetPay += Math.Round(TotalNetPay / 4, 2);
                            AverageHours += Math.Round(TotalHours / 4, 2);
                            TotalEarnings = 0;
                            TotalNetPay = 0;
                            TotalHours = 0;
                            Counter = 0;
                        }
                    }
                    /* Divides the remaining entries that did not make it into the 
                     * monthly average and divides them by the counted entries. 
                     */
                    EarningsRemainder += Math.Round(TotalEarnings / Counter, 2);
                    NetPayRemainder += Math.Round(TotalNetPay / Counter, 2);
                    HoursRemainder += Math.Round(TotalHours / Counter, 2);
                    /* Multiplies the remaining counter amount by .25 and adds that
                     * to the amount of total entries.
                     */
                    intEntryMonths += Counter * .25;
                    break;
                }

                // Closes reader and database connections (always do this after opening)
                reader.Close();
                dbConnection.Close();

                /* Calculates the averages by dividing the earnings from each month
                 * by the total amount of entries in the database. Rounds the result
                 * to the nearest hundreth
                 */
                AverageEarnings = Math.Round((AverageEarnings + EarningsRemainder) / intEntryMonths, 2);
                AverageNetPay = Math.Round((AverageNetPay + NetPayRemainder) / intEntryMonths, 2);
                AverageHours = Math.Round((AverageHours + HoursRemainder) / intEntryMonths, 2);

                // Displays the averaged totals in a text label
                txtAverageEarnings.Text = AverageEarnings.ToString();
                txtAverageNetPay.Text = AverageNetPay.ToString();
                txtAverageHours.Text = AverageHours.ToString();

                /* Calculates the average monthly pay by multiplying the already averaged
                 * Net Pay by 4.345 (the average amount of weeks in each month) and rounding
                 * the result to the nearest hundreth. Does the same for daily pay by dividing
                 * the averaged amount of hours by 7 and multiplying the result by my current wage
                 */
                txtAverageMonthlyPay.Text = Math.Round((AverageNetPay * 4.345), 2).ToString();
                txtAverageDailyPay.Text = Math.Round((AverageHours / 7 * Wage), 2).ToString();

                /* Calculates the average savings monthly and daily. Monthly is the average monthly pay
                 * subracted by the expenses monthly calculated in a constant at the top of the script.
                 * Daily pay is the average Daily Pay subracted by the monthly expenses divided by the 
                 * average number of days in a month. Displays in a text label.
                 */
                txtAverageSavingsMonthly.Text = (double.Parse(txtAverageMonthlyPay.Text) - Expenses).ToString();
                txtAverageSavingsDaily.Text = (double.Parse(txtAverageDailyPay.Text) - Math.Round(Expenses / 30.436875, 2)).ToString();

                // Forgetta Bout it
                //txtMonthsWithoutIncome.Text = Math.Round((1818.57 + 600) / (Expenses), 2).ToString();
            }
            catch (Exception ex)
            {
                // Displays the system error message and location of the error
                MessageBox.Show("Error in UpdateInformation Button: \n" + ex);
            }

        }
        public void SendReport()
        {
            // Send report button
            try
            {
                string ToEmail = "cjbrazeau12@gmail.com";
                string RecipientName = "Finance Project";

                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("Incominator", "cjbrazeau12@gmail.com"));
                message.To.Add(new MailboxAddress(RecipientName, ToEmail));
                message.Subject = "Incominator Report";

                var builder = new BodyBuilder();
                builder.HtmlBody = string.Format(@"<h2>Incominator Report:</h2>
<p> Average Earnings Weekly: <b>$" + AverageEarnings + @"</b>
<p> Average NetPay Weekly: <b>$" + AverageNetPay + @"</b>
<p> Average Hours Weekly: <b>" + AverageHours + @"</b>");
                message.Body = builder.ToMessageBody();

                using (var client = new SmtpClient())
                {
                    // Googles SMTP server, port 587
                    client.Connect("smtp.gmail.com", 587, false);

                    // Send from email login
                    client.Authenticate("IDKTestEmailConnection@gmail.com", "bppqbnbmvccjszso");
                    client.Send(message);
                    client.Disconnect(true);

                    // Shows the user when the email is sent
                    MessageBox.Show("Email Sent!!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error in Send report Button: \n" + ex);
            }
        }
        private void btnSendReport_Click(object sender, EventArgs e)
        {
            // Forces a report to be sent
            // Will add automatic reports at some point
            SendReport();
        }

        private void resetVariables()
        {
            TotalHours = 0;
            TotalEarnings = 0;
            TotalNetPay = 0;
            AverageHours = 0;
            AverageEarnings = 0;
            AverageNetPay = 0;
        }

        private void btnUpdateDisplay_Click(object sender, EventArgs e)
        {
            // Not sure how to delete this without errors :(
        }


    }
}
