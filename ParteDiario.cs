using DevComponents.AdvTree;
using SpreadsheetLight;
using System;
using System.Configuration;
using System.Globalization;
using System.Windows.Forms;

namespace Parte_Diario
{
    public partial class ParteDiario : Form
    {
        private readonly SLDocument _parte;
        private readonly parte part;
        private readonly ParteCD p;
        private string cod;
        private string desc;
        private string mat;

        private string Tp;
        private string SiNo;
        private int fila;

        //private int uidp;//ultimo idParte
        //private int uidd;// ultimo iddetalle
        public ParteDiario()
        {
            InitializeComponent();
            p = new ParteCD();
            part = new parte();
            fila = 3;
            //var directorio = System.IO.Directory.GetCurrentDirectory();
            var appSettings = ConfigurationManager.AppSettings;
            var directorio = appSettings["dirPlantilla"].ToString();
            _parte = new SLDocument(directorio + "\\Plantilla.xlsx");
            //uidp = p.ultimoParte();
            //uidd = p.ultimoDetalle();
            dgvDetalle.Width = this.Width - 20;
            dgvDetalle.Height = this.Height - dgvDetalle.Location.Y - 50;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int i;
            i = 5;
            
            _parte.SelectWorksheet("Listas");

            
            while (_parte.GetCellValueAsString("A" + i) != "")
            {
                lstmovil.Items.Add(_parte.GetCellValueAsString("A" + i));
                i++;
            }
            i = 4;
            while (_parte.GetCellValueAsString("B" + i) != "")
            {
                byte[] utf8=System.Text.Encoding.UTF8.GetBytes(_parte.GetCellValueAsString("B" + i));//luego buscar en la planilla los codigos para guardar
                cod = System.Text.Encoding.UTF8.GetString(utf8);
                desc = _parte.GetCellValueAsString("C" + i);
                
                lstCod1.Items.Add(cod + ":" + desc);
                lstCod2.Items.Add(cod + ":" + desc);
                lstCod3.Items.Add(cod + ":" + desc);
                lstCod4.Items.Add(cod + ":" + desc);
                lstCod5.Items.Add(cod + ":" + desc);
                lstCod6.Items.Add(cod + ":" + desc);

                i++;
            }
            i = 4;
            while (_parte.GetCellValueAsString("D" + i) != "")
            {
                mat = _parte.GetCellValueAsString("D" + i);
                lstMat1.Items.Add(mat);
                lstMat2.Items.Add(mat);
                lstMat3.Items.Add(mat);
                lstMat4.Items.Add(mat);
                lstMat5.Items.Add(mat);
                lstMat6.Items.Add(mat);
                i++;
            }

            for (i = 0; i < 24; i++)
            {
                for (int j = 0; j < 60; j++)
                {
                    if (i < 10)
                    {
                        if (j < 10)
                        {
                            lstHinicio.Items.Add("0" + i + ":0" + j);
                            lstHfin.Items.Add("0" + i + ":0" + j);
                        }
                        else
                        {
                            lstHinicio.Items.Add("0" + i + ":" + j);
                            lstHfin.Items.Add("0" + i + ":" + j);
                        }
                    }
                    else
                    {
                        if (j < 10)
                        {
                            lstHinicio.Items.Add(i + ":0" + j);
                            lstHfin.Items.Add(i + ":0" + j);
                        }
                        else
                        {
                            lstHinicio.Items.Add(i + ":" + j);
                            lstHfin.Items.Add(i + ":" + j);
                        }
                    }
                }
            }
            i = 5;
            while (_parte.GetCellValueAsString("E" + i) != "")
            {
                Tp = _parte.GetCellValueAsString("E" + i);
                lstTipoTrabajo.Items.Add(Tp);
                i++;
            }
            i = 5;
            while (_parte.GetCellValueAsString("F" + i) != "")
            {
                SiNo = _parte.GetCellValueAsString("F" + i);
                lstFinalizado.Items.Add(SiNo);
                i++;
            }
            lstServicio.Items.Add("FTTH");
            lstServicio.Items.Add("WIRELESS");
            lstServicio.Items.Add("DSL");
            btnGuardar.Text = "Guardar";
            btnFinalizar.Text = "Finalizar";
            CargaTecnico();
            txtFecha.Text = DateTime.Today.ToString("dd-MM-yy");
            btnGuardar.Enabled = false;
            btnFinalizar.Enabled = false;
            lstTecnico1.Select();




        }//fin Form_load

        private void lstHinicio_Enter(object sender, EventArgs e)
        {
            string hora;
            int i = 0;
            if (((ListBox)sender).SelectedIndex < 0)
            {
                hora = DateTime.Now.ToString("HH:mm");
                foreach (string item in lstHinicio.Items)
                {
                    if (hora == item.ToString())
                    {
                        lstHinicio.SelectedIndex = i;
                        break;
                    }
                    i++;
                }
            }

            //((ListBox)sender).Enabled = false;

        }

        private void lstHfin_Enter(object sender, EventArgs e)
        {
            string hora;
            int i = 0;
            if (((ListBox)sender).SelectedIndex < 0)
            {
                hora = DateTime.Now.ToString("HH:mm");
                foreach (string item in lstHfin.Items)
                {
                    if (hora == item.ToString())
                    {
                        lstHfin.SelectedIndex = i;
                        break;
                    }
                    i++;
                }
            }

            //((ListBox)sender).Enabled = false;
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            
            cargaDgvDetalle();
            
            btnGuardar.Enabled = false;

        }

        private void btnFinalizar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Quiere finalizar el Parte diario?", "Finalizar", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                var directorio = System.IO.Directory.GetCurrentDirectory();
                var direc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                _parte.SaveAs(direc + "\\parte diario\\parte diario - " + txtFecha.Text + " - " + lstTecnico1.Text + " - " + lstTecnico2.Text + ".xlsx");
                _parte.Dispose();
                this.Close();
            }
            else
            {
                txtTramite.Select();
            }

        }
        /*
         * guardar en planilla
         * hay que modificar para que recorra el datagridview
         */
        private void gaurdarPlanilla()
        {
            string celda;
            string[] tag;//vector que contiene en [0] letra de celda y en [1] numero de casillero del dgvDetalle
            string[] codigo;


            _parte.SelectWorksheet("Parte_Diario");
            foreach (var item in Controls)
            {

                if (item.GetType() == typeof(TextBox))
                {
                    TextBox c = (TextBox)item;
                    tag = c.Tag.ToString().Split(':');
                    celda = tag[0].ToString() + fila.ToString();
                    if (_parte != null)
                        _parte.SetCellValue(celda, c.Text.ToString());
                }

                if (item is ListBox)
                {
                    ListBox l = (ListBox)item;
                    tag = l.Tag.ToString().Split(':');
                    celda = tag[0].ToString() + fila.ToString();
                    if (_parte != null)
                    {
                        if (l.TabIndex >= 16 && l.TabIndex <= 21)
                        {
                            /*
                             * aca aestraigo de una cadena tipo '0.1:trabajo' a un vector donde lo primero es el codigo
                             * y lo segundo es la descripcion
                             */
                            codigo = l.Text.Split(':');
                            _parte.SetCellValue(celda, codigo[0]);

                        }
                        else
                        {
                            _parte.SetCellValue(celda, l.Text);
                        }

                    }
                }

            }
            fila++;
        }


        private void limpiar()
        {
            foreach (var item in Controls)
            {

                if (item.GetType() == typeof(TextBox))
                {
                    TextBox c = (TextBox)item;
                    if (c.TabIndex!=49)
                        c.Text = "";
                }

                if (item is ListBox)
                {
                    ListBox l = (ListBox)item;

                    if (l.Enabled == true)
                    {
                        if(!(l.TabIndex>=50 && l.TabIndex <=52))
                        {
                            if(!(l.TabIndex >= 55 && l.TabIndex <= 56))
                                l.ClearSelected();
                        }
                    }
                        
                }

            }
            txtTramite.Enabled = true;
            txtTramite.Select();
        }


        private void CargaTecnico()
        {

            var appSettings = ConfigurationManager.AppSettings;

            if (appSettings.Count == 0)
            {
                Console.WriteLine("AppSettings is empty.");
            }
            else
            {
                foreach (var key in appSettings.AllKeys)
                {

                    if (key.ToString().Contains("tec"))//asi debo diferenciar las configuragiones
                    {
                        lstTecnico1.Items.Add(appSettings[key].ToString());
                        lstTecnico2.Items.Add(appSettings[key].ToString());
                        lstTecnico3.Items.Add(appSettings[key].ToString());
                        lstTecnico4.Items.Add(appSettings[key].ToString());
                    }

                }
            }

        }

        private void lstmovil_Enter(object sender, EventArgs e)
        {
            int i = 0;

            foreach (string item in lstmovil.Items)
            {

                if ("15" == item.ToString())
                {
                    lstmovil.SelectedIndex = i;
                    break;
                }
                i++;
            }
        }

        private void GuardaNube()
        {
            
            detalle dt = new detalle();

            string[] cod;
            
            //dt.idDesc = uidd + 1;
            dt._idParte = part.idParte;
            dt.servicio = p.buscarServicio(lstServicio.Text).idServicio;
            dt.movil = p.buscarMovil(Convert.ToInt32(lstmovil.Text)).IdMovil;
            dt.horaini = Convert.ToDateTime(lstHinicio.Text);
            dt.horafin = Convert.ToDateTime(lstHfin.Text);
            dt.tipotrab = p.buscarTipoTrabajo(lstTipoTrabajo.Text).IdTipoTrabajo;
            dt.finalizado = lstFinalizado.Text;
            dt.observacion = txtObservacion.Text;
            if (lstCod1.SelectedIndex >= 0)
            {
                cod = lstCod1.Text.Split(':');
                dt.cod1 = p.buscarTarea(cod[0]).idTarea;
            }
            else
            {
                dt.cod1 = 0;
            }
            if (lstCod2.SelectedIndex >= 0)
            {
                cod = lstCod2.Text.Split(':');
                dt.cod2 = p.buscarTarea(cod[0]).idTarea;
            }
            else
            {
                dt.cod2 = 0;
            }
            if (lstCod3.SelectedIndex >= 0)
            {
                cod = lstCod3.Text.Split(':');
                dt.cod3 = p.buscarTarea(cod[0]).idTarea;
            }
            else
            {
                dt.cod3 = 0;
            }
            if (lstCod4.SelectedIndex >= 0)
            {
                cod = lstCod4.Text.Split(':');
                dt.cod4 = p.buscarTarea(cod[0]).idTarea;
            }
            else
            {
                dt.cod4 = 0;
            }
            if (lstCod5.SelectedIndex >= 0)
            {
                cod = lstCod5.Text.Split(':');
                dt.cod5 = p.buscarTarea(cod[0]).idTarea;
            }
            else
            {
                dt.cod5 = 0;
            }
            if (lstCod6.SelectedIndex >= 0)
            {
                cod = lstCod1.Text.Split(':');
                dt.cod6 = p.buscarTarea(cod[0]).idTarea;
            }
            else
            {
                dt.cod6 = 0;
            }
            if (lstMat1.SelectedIndex >= 0)
            {
                dt.mat1 = p.buscarMaterial(lstMat1.Text).idMat;
                dt.cmat1 = Convert.ToInt32(txtCant1.Text);
            }
            else
            {
                dt.mat1 = 0;
                dt.cmat1 = 0;
            }
            if (lstMat2.SelectedIndex >= 0)
            {
                dt.mat2 = p.buscarMaterial(lstMat2.Text).idMat;
                dt.cmat2 = Convert.ToInt32(txtCant2.Text);
            }
            else
            {
                dt.mat2 = 0;
                dt.cmat2 = 0;
            }
            if (lstMat3.SelectedIndex >= 0)
            {
                dt.mat3 = p.buscarMaterial(lstMat3.Text).idMat;
                dt.cmat3 = Convert.ToInt32(txtCant3.Text);
            }
            else
            {
                dt.mat3 = 0;
                dt.cmat3 = 0;
            }
            if (lstMat4.SelectedIndex >= 0)
            {
                dt.mat4 = p.buscarMaterial(lstMat4.Text).idMat;
                dt.cmat4= Convert.ToInt32(txtCant4.Text);
            }
            else
            {
                dt.mat4 = 0;
                dt.cmat4 = 0;
            }
            if (lstMat5.SelectedIndex >= 0)
            {
                dt.mat5 = p.buscarMaterial(lstMat5.Text).idMat;
                dt.cmat5 = Convert.ToInt32(txtCant5.Text);
            }
            else
            {
                dt.mat5 = 0;
                dt.cmat5 = 0;
            }
            if (lstMat6.SelectedIndex >= 0)
            {
                dt.mat6 = p.buscarMaterial(lstMat6.Text).idMat;
                dt.cmat6 = Convert.ToInt32(txtCant6.Text);
            }
            else
            {
                dt.mat6 = 0;
                dt.cmat6 = 0;
            }

            if (lstTecnico1.SelectedIndex >= 0)
            {
                dt.tec1 = lstTecnico1.SelectedIndex;
            }
            else
            {
                dt.tec1 = 0;
            }
            if (lstTecnico2.SelectedIndex >= 0)
            {
                dt.tec2 = lstTecnico2.SelectedIndex;
            }
            else
            {
                dt.tec2 = 0;
            }
            if (lstTecnico3.SelectedIndex >= 0)
            {
                dt.tec3 = lstTecnico3.SelectedIndex;
            }
            else
            {
                dt.tec3 = 0;
            }
            if (lstTecnico4.SelectedIndex >= 0)
            {
                dt.tec4 = lstTecnico4.SelectedIndex;
            }
            else
            {
                dt.tec4 = 0;
            }

            p.altaDetalle(dt);


        }
        private void cargaDgvDetalle()
        {
            DataGridViewRow r = (DataGridViewRow)dgvDetalle.Rows[0].Clone();

            string[] datos;
            string[] codigo;
            foreach (var item in Controls)
            {

                if (item.GetType() == typeof(TextBox))
                {
                    TextBox c = (TextBox)item;


                    if (c.TabIndex >= 0 && c.TabIndex <= 7 || c.TabIndex >= 29 && c.TabIndex <= 39)
                    {
                        datos = c.Tag.ToString().Split(':');
                        r.Cells[Convert.ToInt32(datos[1])].Value = c.Text;
                        r.Cells[Convert.ToInt32(datos[1])].ToolTipText = datos[0];

                    }
                }

                if (item is ListBox)
                {
                    ListBox l = (ListBox)item;

                    if (l.TabIndex >= 1 && l.TabIndex <= 38 || l.TabIndex >= 50 && l.TabIndex <= 52)
                    {
                        datos = l.Tag.ToString().Split(':');
                        codigo = l.Text.Split(':');
                        if (l.TabIndex >= 16 && l.TabIndex <= 21)
                        {
                            r.Cells[Convert.ToInt32(datos[1])].Value = codigo[0];
                            r.Cells[Convert.ToInt32(datos[1])].ToolTipText = codigo[0];
                        }
                        else
                        {

                            if (l.TabIndex == 55 && l.Text != "")
                            {
                                r.Cells[27].Value = r.Cells[27].Value = l.Text + ", " + l.Text;
                                r.Cells[Convert.ToInt32(datos[1])].ToolTipText = codigo[0];
                            }
                            else if (l.TabIndex == 56 && l.Text != "")
                            {
                                r.Cells[27].Value = r.Cells[27].Value = l.Text + ", " + l.Text;
                                r.Cells[Convert.ToInt32(datos[1])].ToolTipText = codigo[0];
                            }
                            else
                            {
                                r.Cells[Convert.ToInt32(datos[1])].Value = l.Text;
                                r.Cells[Convert.ToInt32(datos[1])].ToolTipText = codigo[0];
                            }


                        }


                    }

                }

            }
            dgvDetalle.Rows.Add(r);
            //GuardaNube();
            limpiar();
        }

       

        private void ParteDiario_Resize(object sender, EventArgs e)
        {
            dgvDetalle.Width = ((Form)sender).Width - 20;
            dgvDetalle.Height = ((Form)sender).Height - dgvDetalle.Location.Y - 50;
        }

       
        private void lstServicio_Enter(object sender, EventArgs e)
        {
            /*
            if (MessageBox.Show(txtTramite.Text.ToString(), "Numero de tramite correcto?", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                part.idParte = uidp + 1;
                part.numParte = Convert.ToInt32(txtTramite.Text);
                p.altaParte(part.numParte);
                txtTramite.Enabled = false;

            }
            else
            {
                txtTramite.Select();
            }*/
        }

        private void txtTramite_TextChanged(object sender, EventArgs e)
        {
            btnGuardar.Enabled = true;
            btnFinalizar.Enabled = true;
        }

        private void dgvDetalle_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView grilla = (DataGridView)sender;
            foreach (DataGridViewCell item in grilla.Rows[0].Cells)
            {
                if (item.Selected == true)
                {
                    MessageBox.Show(item.ToolTipText.ToString());
                }
            }
            
        }
    }//fin class
}
