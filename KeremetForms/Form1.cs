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

namespace KeremetForms
{
    public partial class FormMain : Form
    {
        SqlConnection connection;
        string connectionString;

        // Create new Table [Clients]
        string csql = "CREATE TABLE IF NOT EXISTS Clients (" +
                        "ID SERIAL PRIMARY KEY, " +
                        "Name varchar(100) not null, " +
                        "BirthDate DATE not null, " +
                        "PhoneNumber varchar(100), " +
                        "Address varchar(250), " +
                        "SocialNumber varchar(20) not null);";

        string tex = "CREATE TABLE IF NOT EXISTS Derevo(ID SERIAL PRIMARY KEY, NAME VARCHAR(100) NOT NULL, BIRTHDATE DATE NOT NULL, PHONENUMBER VARCHAR(100), ADDRESS VARCHAR(250), SOCIALNUMBER VARCHAR(20) NOT NULL);";

        // Insert some Values to Clients
        string isql = "insert into Clients (Name, BirthDate, SocialNumber) " +
                        "values('Mike', '1985-10-11', @SocialNumber);";

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
            

            for (int i = 250; i < 280; i++)
            {
                scmd.CommandText = "insert into Clients (Name, BirthDate, SocialNumber) " +
                        $"values('Customer-{i.ToString()}', '1985-10-11', '{(12345678901234+i).ToString()}');";
                scmd.ExecuteNonQuery();
            }
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
