using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;

namespace Parte_Diario
{
    public class ParteCD
    {


        public int altaParte(int numParte)
        {
            /*
             * Esta funcion crea un nuevo parte y devuelve el id del mismo
             * 
             */
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();
            int idParte;
            idParte = ultimoParte() + 1;
            try
            {


                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into parte (idParte, numparte) values (" +
                  idParte + ", " + numParte + ")";
                cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return idParte;

        }
        public int altaDetalle(detalle det)
        {
            /*
             * esta funcion sirve para dar de alta el detalle del parte
             * pero debe cargarse todos los datos en el obejo detalle antes
             */
            var sql = "INSERT INTO `detalle` (`idDesc`, `_idParte`, `servicio`, `movil`, `horaini`, `horafin`, `tipotrab`, `finalizado`, `observacion`, `cod1`, `cod2`, `cod3`, `cod4`, `cod6`, `mat1`, `cmat1`, `mat2`, `cmat2`, `mat3`, `cmat3`, `mat4`, `cmat4`, `mat5`, `cmat5`, `mat6`, `cmat6`, `tec1`, `tec2`, `cod5`, `Tec3`, `Tec4`) VALUES (" + det.idDesc + ", " + det._idParte + ", " + det.servicio + ", " + det.movil + ", '" + det.horaini.ToString("HH:mm") + "', '" + det.horafin.ToString("HH:mm") + "', " + det.tipotrab + ", '" + det.finalizado + "', '" + det.observacion + "', " + det.cod1 + ", " + det.cod2 + ", " + det.cod3 + ", " + det.cod4 + ", " + det.cod6 + ", " + det.mat1 + ", " + det.cmat1 + ", " + det.mat2 + ", " + det.cmat2 + ", " + det.mat3 + ", " + det.cmat3 + ", " + det.mat4 + ", " + det.cmat4 + ", " + det.mat5 + ", " + det.cmat5 + ", " + det.mat6 + ", " + det.cmat6 + ", " + det.tec1 + ", " + det.tec2 + ", " + det.cod5 + ", " + det.tec3 + ", " + det.tec4 + ");";
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();
            int rta = 0;
            try
            {


                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = sql;
                rta = cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return rta;
        }

        public movil buscarMovil(int numMovil)
        {
            /*
             * esta funcion devuelve un objeto movil(automotor)
             */
            movil rmovil = new movil();
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();


            try
            {
                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT idMovil, nmovil, patente, marca, anio FROM movil WHERE nmovil =" + numMovil + ";";

                MySqlDataReader rdr = cmd.ExecuteReader();
                rdr.Read();

                rmovil.IdMovil = (int)rdr[0];
                rmovil.nmovil = (int)rdr[1];
                rmovil.patente = (string)rdr[2];
                rmovil.marca = (string)rdr[3];
                rmovil.anio = (int)rdr[4];

            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return rmovil;
        }

        public materiales buscarMaterial(string mat)
        {
            /*
             * esta funcion devuelve un objeto material
             */
            materiales rmat = new materiales();
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();


            try
            {
                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT idMat, material, descripcion FROM materiales WHERE material = '" + mat + "';";

                MySqlDataReader rdr = cmd.ExecuteReader();
                rdr.Read();

                rmat.idMat = (int)rdr[0];
                rmat.material = (string)rdr[1];
                rmat.descripcion = (string)rdr[2];

            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return rmat;
        }

        public servicio buscarServicio(string srv)
        {
            /*
            * esta funcion devuelve un objeto servicio
            */
            servicio rserv = new servicio();
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();


            try
            {
                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT idServicio, Servicio, descripcion FROM servicio WHERE servicio = '" + srv + "';";

                MySqlDataReader rdr = cmd.ExecuteReader();
                rdr.Read();

                rserv.idServicio = (int)rdr[0];
                rserv._servicio = (string)rdr[1];
                rserv.descripcion = (string)rdr[2];

            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return rserv;
        }

        public tarea buscarTarea(string codigoTarea)
        {
            /*
            * esta funcion devuelve un objeto tarea
            */
            tarea rmat = new tarea();
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();


            try
            {
                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT idtarea, codigo, descripcion FROM tarea WHERE codigo = '" + codigoTarea + "';";

                MySqlDataReader rdr = cmd.ExecuteReader();
                rdr.Read();

                rmat.idTarea = (int)rdr[0];
                rmat.codigo = (string)rdr[1];
                rmat.descripcion = (string)rdr[2];

            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return rmat;
        }

        public tipotrabajo buscarTipoTrabajo(string tpt)
        {
            /*
            * esta funcion devuelve un objeto tipotrabajo
            */
            tipotrabajo rmat = new tipotrabajo();
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();


            try
            {
                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT idTipoTrabajo, tipotrabajo, descripcion FROM tipotrabajo WHERE tipotrabajo = '" + tpt + "';";

                MySqlDataReader rdr = cmd.ExecuteReader();
                rdr.Read();

                rmat.IdTipoTrabajo = (int)rdr[0];
                rmat._tipotrabajo = (string)rdr[1];
                rmat.descripcion = (string)rdr[2];

            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return rmat;
        }
        public int ultimoParte()
        {
            /*
             * esta funcion devuelve el ultimo idparte 
             */
            int numparte = 0;
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();


            try
            {
                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT MAX(idParte) as up FROM parte;";

                MySqlDataReader rdr = cmd.ExecuteReader();

                rdr.Read();
                numparte = (int)rdr[0];

            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return numparte;
        }

        public int ultimoDetalle()
        {
            /*
             * esta funcion devuelve el ultimo detalle del parte 
             */
            int idDetalle = 0;
            var conn = new MySqlConnection();
            var cmd = new MySqlCommand();
            var db = new Conectar();



            try
            {
                conn = db.Abrir();
                cmd.Connection = conn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT MAX(iddesc) as up FROM detalle;";

                MySqlDataReader rdr = cmd.ExecuteReader();


                rdr.Read();
                idDetalle = Convert.ToInt32(rdr[0]);



            }
            catch (Exception e)
            {
                MessageBox.Show(String.Concat(e.Message, e.StackTrace), "Control de errores", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return idDetalle;
        }
    }
}
