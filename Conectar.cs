using MySql.Data.MySqlClient;
using System;
using System.Configuration;
using System.Windows.Forms;

namespace Parte_Diario
{

    public class Conectar
    {
        private readonly string _database;

        public Conectar()
        {
            try
            {
                _database = ConfigurationManager.AppSettings["database"];
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "");
            }

        }

        public MySqlConnection Abrir()
        {
            string conexion = ConfigurationManager.ConnectionStrings["conexion4"].ConnectionString;
            var conn = new MySqlConnection();
            conn.ConnectionString = conexion;
            conn.Open();
            return conn;
        }
    }
}
