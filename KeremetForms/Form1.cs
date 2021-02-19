using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
//using System.Data;
using System.Data.SqlClient;
using Npgsql;
using IronXL;

namespace KeremetForms
{
    public partial class FormMain : Form
    {
        SqlConnection connection;
        string connectionString;
        string[] locations = { "Нарын", "Бишкек", "Исфана", "Каракол", "Комсомольское" };

        // Create new Table [Clients]
        string csql = "CREATE TABLE IF NOT EXISTS Clients (" +
                        "ID SERIAL PRIMARY KEY, " +
                        "Name varchar(100) not null, " +
                        "BirthDate DATE not null, " +
                        "PhoneNumber varchar(100), " +
                        "Address varchar(250), " +
                        "SocialNumber varchar(20) not null);";

        public FormMain()
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["KeremetForms.Properties.Settings.KeremetDBConnectionString"].ConnectionString;

            //InitializeDb(connectionString);
            PsqlDbInitialize();
        }

        private void PsqlDbInitialize()
        {
            var cs = "Host=localhost;Username=postgres;Password=doomSpawnMk;Database=vstest";
            NpgsqlConnection con = new NpgsqlConnection(cs);
            con.Open();

            NpgsqlCommand scmd = new NpgsqlCommand();
            scmd.Connection = con;

            scmd.CommandText = "DROP TABLE IF EXISTS Clients";
            scmd.ExecuteNonQuery();

            scmd.CommandText = csql; // Create Table in PG DB
            scmd.ExecuteNonQuery();

            Random rand = new Random();

            for (int i = 250; i < 280; i++)
            {
                var rd = RandomDayFunc();
                string birtdate = $"'{rd().Year}-{rd().Month}-{rd().Day}'";
                string phone = $"'{rand.Next(100, 790)}'";
                string socialnumber = $"'{12345678901234 + i}'";
                string address = $"'{locations[rand.Next(locations.Length)]}'";

                scmd.CommandText = "insert into Clients (Name, BirthDate, SocialNumber, PhoneNumber, Address) " +
                        $"values('Customer-{i}', {birtdate}, {socialnumber}, {phone}, {address});";
                scmd.ExecuteNonQuery();
            }

            Func<DateTime> RandomDayFunc() // Making random Date of birth
            {
                DateTime start = new DateTime(1995, 1, 1);
                Random gen = new Random();
                int range = ((TimeSpan)(new DateTime(2000,12,31) - start)).Days;
                return () => start.AddDays(gen.Next(range));
            }
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            string socNum = txtInput.Text;

            string cs = "Host=localhost;Username=postgres;Password=doomSpawnMk;Database=vstest";
            NpgsqlConnection con = new NpgsqlConnection(cs);
            con.Open();

            NpgsqlCommand scmd = new NpgsqlCommand();
            scmd.Connection = con;

            scmd.CommandText = $"SELECT * FROM public.clients WHERE socialnumber = '{socNum}'";
            NpgsqlDataReader client = scmd.ExecuteReader();

            var workbook = new WorkBook("D:/coding/C#/KeremetForms/KeremetForms/Template/example.xlsx");
            var worksheet = workbook.GetWorkSheet("Лист1");
            var range = worksheet.GetRange("A1:Z20");

            while (client.Read())
            {
                foreach (var cell in range)
                {
                    if (cell.Value.ToString() == "[Name]")
                    {
                        cell.Value = client["Name"].ToString();
                    }
                    else if (cell.Value.ToString() == "[BirthDate]")
                    {
                        cell.Value = client["BirthDate"].ToString();
                    }
                    else if (cell.Value.ToString() == "[SocialNumber]")
                    {
                        cell.Value = client["SocialNumber"].ToString();
                    }
                    else if (cell.Value.ToString() == "[PhoneNumber]")
                    {
                        cell.Value = client["PhoneNumber"].ToString();
                    }
                    else if (cell.Value.ToString() == "[Address]")
                    {
                        cell.Value = client["Address"].ToString();
                    }
                }
            }

            workbook.SaveAs($"D:/coding/C#/KeremetForms/KeremetForms/Result/client_{txtInput.Text}.xlsx")
            con.Close();
        }

        private void InitializeDb(string connectionString)
        {
            // Create new Table [Clients]
            string csql = "CREATE TABLE IF NOT EXISTS Clients (" +
                            "ID INT PRIMARY KEY, " +
                            "Name varchar(100) not null, " +
                            "BirthDate DATE not null, " +
                            "PhoneNumber varchar(100), " +
                            "Address varchar(250), " +
                            "SocialNumber varchar(20) not null);";
            
            string tex = "CREATE TABLE IF NOT EXISTS Derevo(ID SERIAL PRIMARY KEY, NAME VARCHAR(100) NOT NULL, BIRTHDATE DATE NOT NULL, PHONENUMBER VARCHAR(100), ADDRESS VARCHAR(250), SOCIALNUMBER VARCHAR(20) NOT NULL);";
            
            // Insert some Values to Clients
            string isql = "insert into Clients(Name, BirthDate, SocialNumber) " +
                            "values('Mike', '1985-10-11', '@SocialNumber');";

            using (connection = new SqlConnection(connectionString))
            using (SqlCommand command = new SqlCommand())
            {
                connection.Open();
                command.Connection = connection;
                command.CommandText = tex;
                command.ExecuteScalar();

                for (var i = 120245680; i < 120245690; i++)
                {
                    command.CommandText = isql;
                    command.Parameters.AddWithValue("@SocialNumber", i.ToString());
                    command.ExecuteNonQuery();
                }
            }
        }

    }
}
