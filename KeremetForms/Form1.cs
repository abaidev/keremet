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
        string cs = "Host=localhost;Username=postgres;Password=doomSpawnMk;Database=vstest";
        string[] locations = { "Нарын", "Бишкек", "Исфана", "Каракол", "Комсомольское" };
        NpgsqlConnection con;
        NpgsqlCommand scmd;

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
            
            PsqlDbInitialize();
        }

        private void PsqlDbInitialize()
        {
            con = new NpgsqlConnection(cs);
            con.Open();

            scmd = new NpgsqlCommand();
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
                int range = ((TimeSpan)(new DateTime(2000, 12, 31) - start)).Days;
                return () => start.AddDays(gen.Next(range - 10));
            }

            con.Close();
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            string socNum = txtInput.Text;

            con = new NpgsqlConnection(cs);
            con.Open();

            scmd = new NpgsqlCommand();
            scmd.Connection = con;

            var workbook = new WorkBook("D:/coding/C#/KeremetForms/KeremetForms/Template/example.xlsx");
            var worksheet = workbook.GetWorkSheet("Лист1");
            var range = worksheet.GetRange("A1:Z20");

            try
            {
                //scmd.CommandText = $"SELECT * FROM Clients WHERE SocialNumber = '{socNum}'"; // also works fine, but a bit slower
                scmd.CommandText = $"SELECT * FROM public.clients WHERE socialnumber = '{socNum}'"; // with this finds very quickly. i guess because of PG
                NpgsqlDataReader client = scmd.ExecuteReader();

                if (client.Read())
                {
                    lblFound.Text = $"{client["Name"]}";
                    
                    DialogResult result = MessageBox.Show($"Сохранить клиента {client["Name"]} в файл .xlsx?", "Подтверждение", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        foreach (var cell in range)
                        {
                            if (cell.Value.ToString() == "[ID]")
                            {
                                cell.Value = client["Id"].ToString();
                            }
                            if (cell.Value.ToString() == "[Name]")
                            {
                                cell.Value = client["Name"].ToString();
                            }
                            else if (cell.Value.ToString() == "[BirthDate]")
                            {
                                cell.Value = String.Format("{0:dd-MM-yyyy}", client["BirthDate"]); ;
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
                        workbook.SaveAs($"D:/coding/C#/KeremetForms/KeremetForms/Result/client_{txtInput.Text}.xlsx");
                        MessageBox.Show("Клиент успешно сохранен в папке Result.", "status", MessageBoxButtons.OK);
                    }
                    else
                    {
                        MessageBox.Show("Вы отменили сохранение.", "status", MessageBoxButtons.OK);
                    }
                }
                else
                {
                    lblFound.Text = "отсутствует";
                    MessageBox.Show("Неверный ИНН. Попробуйте еще раз");
                }
            }
            catch(Exception exc)
            {
                MessageBox.Show(exc.Message);
            }                    
            con.Close();
        }

    }
}
