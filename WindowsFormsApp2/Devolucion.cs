using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using XSDToXML.Utils;
using System.Diagnostics;
using CryptoSysPKI;
using System.Net;
using System.Net.Mail;
using System.Configuration;

namespace WindowsFormsApp2
{
    public partial class Devolucion : Form
    {
        Main serv = new Main();

        //Direcciones para guardar los XML
        static string pathXML;
        static private string pathxGDL = @"\\192.168.0.2\Sistemas\Sistemas\CFD\Factura3.3\Factura3.3\bin\Debug\Factura\";
        static private string pathxATQ = @"\\192.168.1.2\Sistemas\Sistemas\CFD\Factura3.3\Factura3.3\bin\Debug\Factura\";
        static private string pathxIXT = @"\\192.168.2.2\Sistemas\Sistemas\CFD\Factura3.3\Factura3.3\bin\Debug\Factura\";
        static private string pathxTLA = @"\\192.168.3.2\Sistemas\Sistemas\CFD\Factura3.3\Factura3.3\bin\Debug\Factura\";
        /*
        //REALES
        static string pathCer = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\00001000000203050906.cer";
        static string pathKey = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\CRE840310TC6_1302151633S.key";
        static string pathPem = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\Archivo.pem";
        static string claverPrivada = "COM361490";

        //REALES
        String UsuarioFell = "CRE840310TC6";
        String ContraseñaFEll = "GuRgWjw$";
        */
        
        //PRUEBA
        static string pathCer = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\CertificadoFirmadoPM.cer";
        static string pathKey = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\LlavePkcs8PM.key";
        static string pathPem = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\Archivo.pem";
        static string claverPrivada = "12345678a";
      
        //PRUEBA
        String UsuarioFell = "CRE840310D33";
        String ContraseñaFEll = "contRa$3na";
        

        
        decimal totald, descuentod, subtotald, ivaTotal, importeDevolucion;
        decimal totalNegativo, descuentoNegativo, subtotalNegativo;
        int formaPago, opServer, facturaOK = 0, noRefactura = 0;
        string diaCarpeta, pathxDinamico;
        Double importe, Descuento, porcentaje, Subtotal, IVA, Total, auxformapago;
        string claveProdServ, numOpPathXMLFac, numOpPathXMLNota, MetodoPago, numOpSinCambios, importeGlobal;
        string from = "credito@casaguerrero.com.mx", to, pass = "credguerrero", subject = "Comercial del Retiro SA de CV", body = "";
        string operacion, fecha, operaci, cuenta, RFC, nombre, email, telefono, UUIDRELACIONADO, lugarExpedicion, sucursal, numOpCorreo;
        string UUIDCANCELADO, UUID, numOp, numOpProd, totalaux, totalletras, observaciones, numOpNC, numOpFac, usoCFDII, usoCFDI, UUIDNUEVO;
        string servidor, RFCAnterior, formaPagoStr;
        bool tieneAbono = false;


        private void pictureBox5_Click(object sender, EventArgs e)
        {
            this.BackColor = Color.LightGray;
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.BackColor = Color.MintCream;
        }

        private void Devolucion_Load_1(object sender, EventArgs e)
        {

            checkG01.Checked = true;
            comboOtro.Items.Add("G02");
            comboOtro.Items.Add("I02");
            comboOtro.Items.Add("I04");
            comboOtro.Items.Add("I01");
            comboOtro.Items.Add("I03");
            comboOtro.Items.Add("I05");
            comboOtro.Items.Add("I06");
            comboOtro.Items.Add("I07");
            comboOtro.Items.Add("I08");
            comboOtro.Items.Add("D01");
            comboOtro.Items.Add("D02");
            comboOtro.Items.Add("D03");
            comboOtro.Items.Add("D04");
            comboOtro.Items.Add("D05");
            comboOtro.Items.Add("D06");
            comboOtro.Items.Add("D07");
            comboOtro.Items.Add("D08");
            comboOtro.Items.Add("D09");
            comboOtro.Items.Add("D10");

            textNumOp.Enabled = false;
            textCuenta.MaxLength = 10;
            textRFC.MaxLength = 13;
            textNombre.MaxLength = 150;
            textEmail.MaxLength = 50;
            textTelefono.MaxLength = 15;

        }

        private void Facturar_Click(object sender, EventArgs e)
        {
            //1.200.1
            //#################----------------------------------------------------- EN PROGRESO ---------------------------------------------
            Cursor.Current = Cursors.WaitCursor;

            //Datos
            operaci = textOperacion.Text;
            cuenta = textCuenta.Text;
            RFC = textRFC.Text;
            nombre = textNombre.Text;
            email = textEmail.Text;
            telefono = textTelefono.Text;

            if (RFC is null || nombre is null || email is null || RFC == ""|| nombre == "" || email =="")
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("NECESITA LLENAR LOS DATOS RFC, NOMBRE Y EMAIL PARA FACTURAR ESTA OPERACIÓN", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (UUID != "" && UUID != null)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("YA EXISTE UNA FACTURA CON ESTE NÚMERO DE OPERACIÓN", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ButtonFacturar.Enabled = false;
            }
            else
            {
                //Datos necesarios
                ButtonNota.Enabled = false;
                buttonReFac.Enabled = false;
                ButtonFacturar.Enabled = false;
                timbreCreditoDataGridView.Rows[0].Selected = true;
                //MOVTOS ############
                DataGridViewSelectedRowCollection row = timbreCreditoDataGridView.SelectedRows;

                if (timbreCreditoDataGridView.CurrentRow == null)
                {
                    //Uso CFDI
                    if (checkOtro.Checked)
                    {
                        usoCFDII = comboOtro.SelectedItem.ToString();
                    }
                    else
                    {
                        if (checkG01.Checked)
                        {
                            usoCFDII = "G01";
                        }
                        if (checkG03.Checked)
                        {
                            usoCFDII = "G03";
                        }
                        if (checkP01.Checked)
                        {
                            usoCFDII = "P01";
                        }
                    }
                    //Fecha
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = "yyyy-MM-dd H:mm:ss";
                    fecha = dateTimePicker1.Text;

                    //Conexion sql
                    String servidor = GetServidor(opServer);
                    SqlConnection cn = new SqlConnection(servidor);
                    cn.Open();
                    SqlCommand cmd = cn.CreateCommand();

                    numOp = encontrarConsecutivo(numOpFac, operaci, "F");
                    
                    numOp = numOp.Replace("  ", " ");

                    //FACTURA AL SQL ################################################################# NO tiene uuid relacionado, solo va el nuevo
                    cmd.CommandText = "INSERT INTO dbo.Facturas([Numero],[Fecha],[Cuenta],[Nombre],[Domicilio],[Interior],[Exterior],[Colonia],[Ciudad],[Estado],[CP],[Correo],[Telefono],[RFC],[Operacion],[Importe],[Descuento],[Porcentaje],[Subtotal],[IVA],[Total],[Letra],[TotalCFDI],[UsoCFDI],[FacturaNoGenerica],[UUID],[Observaciones],[UUIDRelacionado],[UUIDCancelado]) VALUES('" + numOp + "',@Fech,'" + cuenta + "','" + nombre + "',null,null,null,null,null,null,null,'" + email + "','" + telefono + "','" + RFC + "','" + operaci + "','" + importe + "','" + Descuento + "','" + porcentaje + "','" + Subtotal + "','" + IVA + "','" + Total + "','" + totalletras + "',null,'" + usoCFDII + "',null,null,null,null,null);";
                    cmd.Parameters.Add("@Fech", SqlDbType.Date).Value = dateTimePicker1.Value;
                    cmd.ExecuteNonQuery();

                    timbrarFactura();


                    if (facturaOK == 0)
                    {
                        //Actualizar SQL #################################################################
                        cmd.CommandText = "update dbo.Facturas set UUID = '" + UUIDNUEVO + "', FormaPago ='" + formaPagoStr + "', MetodoPago = '" + MetodoPago.ToString() + "', Observaciones= '" + observaciones + "' where Numero = '" + numOp + "' ";
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        //En caso de que no se haga la factura se borra el registro
                        cmd.CommandText = "delete from dbo.Facturas where Numero = '" + numOp + "'";
                        cmd.ExecuteNonQuery();
                    }


                    if (facturaOK == 0)
                    {
                        //Mover a relacionado#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
                        cmd.CommandText = "update dbo.Facturas set UUIDCANCELADO= '" + UUIDRELACIONADO + "', Numero = '"+ numOpSinCambios+"FC"+ "' where Numero = '" + numOpSinCambios + "' ";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "update dbo.Facturas set Numero = '" + numOpSinCambios+ "' where Numero = '" + numOp + "' ";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "DELETE FROM dbo.Facturas where Numero = '" + numOpSinCambios+"FC"+"' ";
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("NO SE PUEDE REFACTURAR POR QUE \nTIENE ABONOS O ENGANCHE YA FATURADOS", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Cursor.Current = Cursors.Default;
            }
                
        }

        private void Cancelar_Click(object sender, EventArgs e)
        {
            //0.900.2
            //Hacer cancelacion de cualquier factura
            /*
              Public Class Form1

                Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
                'Se instancia el WS de Timbrado.
                Dim ServicioTimbrado_FEL As New WSFELdemo.WSCFDI33Client
                Dim RespuestaServicio_FEL As New WSFELdemo.RespuestaCancelacion

                Dim listaArreglos As List(Of WSFELdemo.DetalleCFDICancelacion) = New List(Of WSFELdemo.DetalleCFDICancelacion)



                Dim RespuestaCancelacionDetallada_FEL As New List(Of CancelarCFDIConValidacion.WSFELdemo.DetalleCancelacion)

                Dim i As Integer
                For i = 0 To 3
                Dim Arreglo As WSFELdemo.DetalleCFDICancelacion = New WSFELdemo.DetalleCFDICancelacion

                Arreglo.RFCReceptor = "TES030201001" + i.ToString
                Arreglo.Total = "1571.43" + i.ToString
                Arreglo.UUID = "34999FE8-7E57-7E57-7E57-7DE3AD8F6F1B" + i.ToString

                listaArreglos.Add(Arreglo)

                Next

                RespuestaServicio_FEL = ServicioTimbrado_FEL.CancelarCFDIConValidacion("CFDI010233123", "Prueba$", "TES030201001", listaArreglos.ToArray, "Key", "12345678a")

                For Each UUID As WSFELdemo.DetalleCancelacion In RespuestaServicio_FEL.DetallesCancelacion

                TextBox1.Text += RespuestaServicio_FEL.OperacionExitosa.ToString + vbNewLine
                TextBox1.Text += RespuestaServicio_FEL.MensajeError + vbNewLine
                TextBox1.Text += RespuestaServicio_FEL.MensajeErrorDetallado + vbNewLine
                TextBox1.Text += RespuestaServicio_FEL.XMLAcuse + vbNewLine
                TextBox1.Text += UUID.CodigoResultado + vbNewLine
                TextBox1.Text += UUID.MensajeResultado + vbNewLine
                TextBox1.Text += UUID.UUID + vbNewLine
                Next
                End Sub
                End Class
            

            var result = MessageBox.Show("¿ESTA SEGURO DE QUE DESEA CANCELAR ESTA FACTURA?", "Cancelar Factura",
                                         MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                ServiceReferenceFac.WSCFDI33Client ServicioTimbrado_FEL = new ServiceReferenceFac.WSCFDI33Client();
                ServiceReferenceFac.RespuestaCancelacion RespuestaServicio_FEL = new ServiceReferenceFac.RespuestaCancelacion();
                List<ServiceReferenceFac.DetalleCancelacion> RespuestaCancelacionDetallada_FEL = new List<ServiceReferenceFac.DetalleCancelacion>();
                ServiceReferenceFac.DetalleCFDICancelacion[] CFDICancelar = new ServiceReferenceFac.DetalleCFDICancelacion[1];
                ServiceReferenceFac.DetalleCFDICancelacion detalleCFDICancelacion = new ServiceReferenceFac.DetalleCFDICancelacion();

                detalleCFDICancelacion.RFCReceptor = "ZUMS841206CN5";
                detalleCFDICancelacion.Total = 213;
                detalleCFDICancelacion.UUID = "EE8C29F3-7E57-7E57-7E57-3CD698DC46A8";

                CFDICancelar[0] = detalleCFDICancelacion;

                // Error de pfx, hacer pruebas y ver resultado de los sellos
                RespuestaServicio_FEL = ServicioTimbrado_FEL.CancelarCFDIConValidacion(UsuarioFell, ContraseñaFEll, "CRE840310TC6", CFDICancelar, pathPem, claverPrivada);
                if (RespuestaServicio_FEL.OperacionExitosa == true)
                {
                    RespuestaCancelacionDetallada_FEL = RespuestaServicio_FEL.DetallesCancelacion.ToList();
                    foreach (ServiceReferenceFac.DetalleCancelacion UUID in RespuestaCancelacionDetallada_FEL)
                    {
                        MessageBox.Show(""+ UUID.CodigoResultado + System.Environment.NewLine + "" + UUID.MensajeResultado + System.Environment.NewLine +""+ UUID.UUID + System.Environment.NewLine);
                    }
                    //Se guarda localmente el acuse de cancelación que contendra el registro de todas las operaciones que realizamos, osea 1 acuse
                    XmlDocument AcuseXML = new XmlDocument();
                    AcuseXML.LoadXml(RespuestaServicio_FEL.XMLAcuse);
                    AcuseXML.Save("C:\\XML\\AcuseCancelacion.xml");
                    MessageBox.Show("LA FACTURA SE CANCELÓ", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Error: " + RespuestaServicio_FEL.MensajeError + System.Environment.NewLine+ "Detalles: " +  RespuestaServicio_FEL.MensajeErrorDetallado);
                    
                    //## ERROR NULO
                    RespuestaCancelacionDetallada_FEL = RespuestaServicio_FEL.DetallesCancelacion.ToList();

                    foreach (ServiceReferenceFac.DetalleCancelacion UUID in RespuestaCancelacionDetallada_FEL)
                    {
                        MessageBox.Show( UUID.CodigoResultado +  UUID.MensajeResultado + UUID.UUID );
                    }
                }
                    //---------------------------------------------------------
            }
            else
            {
                MessageBox.Show("LA FACTURA NO SE CANCELÓ", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            */
        }

      

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            this.BackColor = Color.Ivory;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.BackColor = Color.White;
        }

        //Conexion al servidor app.config
        public String GetServidor(int op)
        {
            switch (op)
            {
                case 1:
                    try
                    {
                        servidor = ConfigurationManager.ConnectionStrings["WindowsFormsApp2.Properties.Settings.dbSIAConexion"].ConnectionString.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR DE CONEXIÓN: "+ex.Message, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                        break;
                case 2:
                    try
                    {
                        servidor = ConfigurationManager.ConnectionStrings["WindowsFormsApp2.Properties.Settings.dbSIAConexionATQ"].ConnectionString.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR DE CONEXIÓN: " + ex.Message, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                        break;
                case 3:
                    try
                    {
                        servidor = ConfigurationManager.ConnectionStrings["WindowsFormsApp2.Properties.Settings.dbSIAConexionIXT"].ConnectionString.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR DE CONEXIÓN: " + ex.Message, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                        break;
                case 4:
                    try
                    {
                        servidor = ConfigurationManager.ConnectionStrings["WindowsFormsApp2.Properties.Settings.dbSIAConexionTLA"].ConnectionString.ToString();
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("ERROR DE CONEXIÓN: " + ex.Message, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                        break;
            }
            return servidor;
        }

        //ESPACIOS
        private void textEmail_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = e.KeyChar == Convert.ToChar(Keys.Space);
        }
        
        private void textTelefono_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = e.KeyChar == Convert.ToChar(Keys.Space);
        }

        private void textRFC_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = e.KeyChar == Convert.ToChar(Keys.Space);
        }
        
        private void textOperacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = e.KeyChar == Convert.ToChar(Keys.Space);
        }


        //SELLAR
        static string Test_Mex_Sign_Data(string cadenaOriginal)
        {
            //CÓDIGO MAGICO, NO LO TOQUES
            string strData;
            StringBuilder sbPassword;
            StringBuilder sbPrivateKey;
            string sello;
            byte[] b;
            byte[] block;
            int keyBytes;

            string strKeyFile = pathKey;
            sbPassword = new StringBuilder(claverPrivada);
            //StringBuilder no string
            sbPrivateKey = Rsa.ReadEncPrivateKey(strKeyFile, sbPassword.ToString());
            Debug.Assert(sbPrivateKey.Length > 0);
            keyBytes = Rsa.KeyBytes(sbPrivateKey.ToString());
            Debug.Assert(keyBytes > 0);
            strData = cadenaOriginal; 
            //strData = "||2.0|A|1|2009-08-16T16:30:00|1|2009|ingreso|Una sola exhibición|350.00|5.25|397.25|ISP900909Q88|Industrias del Sur Poniente, S.A. de C.V.|Alvaro Obregón|37|3|Col. Roma Norte|México|Cuauhtémoc|Distrito Federal|México|06700|Pino Suarez|23|Centro|Monterrey|Monterrey|Nuevo Léon|México|95460|CAUR390312S87|Rosa María Calderón Uriegas|Topochico|52|Jardines del Valle|Monterrey|Monterrey|Nuevo León|México|95465|10|Caja|Vasos decorados|20.00|200|1|pieza|Charola metálica|150.00|150|IVA|15.00|52.50||";
            b = System.Text.Encoding.UTF8.GetBytes(strData);
            block = Rsa.EncodeMsgForSignature(keyBytes, b, HashAlgorithm.Sha256);
            Debug.Assert(block.Length > 0);
            block = Rsa.RawPrivate(block, sbPrivateKey.ToString());
            Debug.Assert(block.Length > 0);
            sello = System.Convert.ToBase64String(block);
            Wipe.String(sbPassword);
            Wipe.String(sbPrivateKey);
            Wipe.Data(block);
            

            return sello;
        }

        //CAMBIA CORREO 
        private void buttonCorreo_Click(object sender, EventArgs e)
        {
            //Cargar datos de seleccionados
            ButtonNota.Enabled = false;
            buttonReFac.Enabled = false;
            ButtonFacturar.Enabled = false;

            DataGridViewSelectedRowCollection row = facturasDataGridView.SelectedRows;
            numOpCorreo = row[0].Cells[4].Value.ToString();
            diaCarpeta = row[0].Cells[5].Value.ToString();
            string uuidNom = row[0].Cells[16].Value.ToString();
            string diaCarpe;
            diaCarpe = diaCarpeta.Substring(0, 2) + "-" + diaCarpeta.Substring(3, 2) + "-" + diaCarpeta.Substring(6, 4) + @"\";
            
            string correoNuevo;
            correoNuevo = textEmail.Text;
            if (uuidNom != "XAXX010101000")
            {
                //Conexion sql
                String servidor = GetServidor(opServer);
                SqlConnection cn = new SqlConnection(servidor);
                cn.Open();
                try
                {
                    SqlCommand cmd = cn.CreateCommand();
                    cmd.CommandText = "update dbo.Facturas set Correo = '" + correoNuevo + "' where Numero = '" + numOpCorreo + "' ";
                    cmd.ExecuteNonQuery();
                    labelInfo.Text = "CORREO ACTUALIZADO CON ÉXITO";
                    EnviarCorreo(pathxDinamico + diaCarpe + numOpCorreo + ".pdf", pathxDinamico + diaCarpe + numOpCorreo + ".xml", correoNuevo);
                }
                catch (Exception errorCo)
                {
                    MessageBox.Show("ERROR: " + errorCo, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("NO SE PUEDE ENVIAR UNA FACTURA CON RFC GENERICO", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //SELECCIONA PARA LLENAR DATOS PARA CAMBIO DE CORREO
        private void facturasDataGridView_MouseClick(object sender, MouseEventArgs e)
        {
            //Cargar datos de seleccionados
            ButtonNota.Enabled = false;
            buttonReFac.Enabled = false;
            ButtonFacturar.Enabled = false;
            
            DataGridViewSelectedRowCollection row = facturasDataGridView.SelectedRows;
            textEmail.Text = row[0].Cells[14].Value.ToString();
            numOpCorreo = row[0].Cells[4].Value.ToString();
        }

        //ENVIAR CORREO 
        private void EnviarCorreo(string documentopdf_, string documentoxml_, string destinatario_)
        {
            
            //to = destinatario_;
            to = "sistemas2comret@outlook.com";
            using (SmtpClient protocoll = new SmtpClient("mail.casaguerrero.com.mx", 2525))
            {
                MailMessage mensaje = new MailMessage(from, to, subject, body);
                try
                {
                    Attachment data = new Attachment(documentopdf_);
                    mensaje.Attachments.Add(data);
                }
                catch { }
                try
                {
                    Attachment data2 = new Attachment(documentoxml_);
                    mensaje.Attachments.Add(data2);
                }
                catch { }
                protocoll.Credentials = new NetworkCredential(from, pass);
                protocoll.EnableSsl = false;
                try
                {
                    protocoll.Send(mensaje);
                    MessageBox.Show("EL CORREO SE ENVIÓ CON ÉXITO", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                }
                catch (Exception errormsj)
                {
                    MessageBox.Show("EL CORREO NO SE ENVIÓ DEBIDO AL ERROR: " + errormsj, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //CANCELACION POR DEVOLUCION
        private void timbrarNotaCredNegativos()
        {
            //Crear XML de nota de credito
            Comprobante oComprobante = new Comprobante();

            oComprobante.Version = "3.3";
            oComprobante.Folio = numOp;//Numero de operacion 

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy-MM-ddTHH:mm:ss";
            oComprobante.Fecha = Convert.ToDateTime(dateTimePicker1.Text.ToString()); //Formato que debe llevar 
            // CASE PARA FORMA DE PAGO

            switch (formaPago)
            {
                case (99):
                    oComprobante.FormaPago = c_FormaPago.Item99;
                    break;
                case (01):
                    oComprobante.FormaPago = c_FormaPago.Item01;
                    break;
                case (04):
                    oComprobante.FormaPago = c_FormaPago.Item04;
                    break;
                case (28):
                    oComprobante.FormaPago = c_FormaPago.Item28;
                    break;
            }
            oComprobante.NoCertificado = ObtenerCertificado();
            oComprobante.SubTotal = Math.Round((totalNegativo / 1.16m), 2);
            if (descuentod == 0)
            {
            }
            else
            {
                oComprobante.Descuento = descuentod;
            }
            oComprobante.Moneda = c_Moneda.MXN;
            oComprobante.Total = Math.Round(totalNegativo, 2);
            oComprobante.TipoDeComprobante = c_TipoDeComprobante.E; //Siempre egreso en nota de credito
            if (MetodoPago == "PUE")
            {
                oComprobante.MetodoPago = c_MetodoPago.PUE;
            }
            else
            {
                oComprobante.MetodoPago = c_MetodoPago.PPD;
            }
            oComprobante.LugarExpedicion = lugarExpedicion;

            ComprobanteCfdiRelacionados oComprobanteCFDI = new ComprobanteCfdiRelacionados();

            oComprobanteCFDI.TipoRelacion = c_TipoRelacion.Item03; // Siempre va este

            ComprobanteCfdiRelacionadosCfdiRelacionado oComprobanteCFDIRel = new ComprobanteCfdiRelacionadosCfdiRelacionado();
            List<ComprobanteCfdiRelacionadosCfdiRelacionado> lstCfdiRelacRelac = new List<ComprobanteCfdiRelacionadosCfdiRelacionado>();

            oComprobanteCFDIRel.UUID = UUIDRELACIONADO;
            lstCfdiRelacRelac.Add(oComprobanteCFDIRel);
            oComprobanteCFDI.CfdiRelacionado = lstCfdiRelacRelac.ToArray();
            oComprobante.CfdiRelacionados = oComprobanteCFDI;


            ComprobanteEmisor oEmisor = new ComprobanteEmisor();
            oEmisor.RegimenFiscal = c_RegimenFiscal.Item601; //Siempre va este
            ComprobanteReceptor oReceptor = new ComprobanteReceptor();
            /*
            //REALES

            oEmisor.Rfc = "CRE840310TC6";
            oEmisor.Nombre = "COMERCIAL DEL RETIRO SA DE CV";
            //REALES
            oReceptor.Nombre = nombre;
            oReceptor.Rfc = RFC;
            */
           
            //PRUEBA
            
            oEmisor.Rfc = "TES030201001";
            oEmisor.Nombre = "temporal";
            
            //PRUEBA
            oReceptor.Nombre = "temporal";
            oReceptor.Rfc = "TEST010203001";
            

            //Asignar emisor y receptor
            oComprobante.Emisor = oEmisor;
            oComprobante.Receptor = oReceptor;

            //Switch uso CFDI
            switch (usoCFDI)
            {
                case "P01":
                    oReceptor.UsoCFDI = c_UsoCFDI.P01;
                    break;
                case "G01":
                    oReceptor.UsoCFDI = c_UsoCFDI.G01;
                    break;
                case "G02":
                    oReceptor.UsoCFDI = c_UsoCFDI.G02;
                    break;
                case "G03":
                    oReceptor.UsoCFDI = c_UsoCFDI.G03;
                    break;
                case "I01":
                    oReceptor.UsoCFDI = c_UsoCFDI.I01;
                    break;
                case "I02":
                    oReceptor.UsoCFDI = c_UsoCFDI.I02;
                    break;
                case "I03":
                    oReceptor.UsoCFDI = c_UsoCFDI.I03;
                    break;
                case "I04":
                    oReceptor.UsoCFDI = c_UsoCFDI.I04;
                    break;
                case "I05":
                    oReceptor.UsoCFDI = c_UsoCFDI.I05;
                    break;
                case "I06":
                    oReceptor.UsoCFDI = c_UsoCFDI.I06;
                    break;
                case "I07":
                    oReceptor.UsoCFDI = c_UsoCFDI.I07;
                    break;
                case "I08":
                    oReceptor.UsoCFDI = c_UsoCFDI.I08;
                    break;
                case "D01":
                    oReceptor.UsoCFDI = c_UsoCFDI.D01;
                    break;
                case "D02":
                    oReceptor.UsoCFDI = c_UsoCFDI.D02;
                    break;
                case "D03":
                    oReceptor.UsoCFDI = c_UsoCFDI.D03;
                    break;
                case "D04":
                    oReceptor.UsoCFDI = c_UsoCFDI.D04;
                    break;
                case "D05":
                    oReceptor.UsoCFDI = c_UsoCFDI.D05;
                    break;
                case "D06":
                    oReceptor.UsoCFDI = c_UsoCFDI.D06;
                    break;
                case "D07":
                    oReceptor.UsoCFDI = c_UsoCFDI.D07;
                    break;
                case "D08":
                    oReceptor.UsoCFDI = c_UsoCFDI.D08;
                    break;
                case "D09":
                    oReceptor.UsoCFDI = c_UsoCFDI.D09;
                    break;
                case "D10":
                    oReceptor.UsoCFDI = c_UsoCFDI.D10;
                    break;

            }

            decimal totaldd = totalNegativo - (totalNegativo * 0.16m);

            List<ComprobanteConcepto> lstConceptos = new List<ComprobanteConcepto>();
            List<ComprobanteConceptoImpuestos> lstimpuesto = new List<ComprobanteConceptoImpuestos>();
            List<ComprobanteConceptoImpuestosTraslado> lstTrasla = new List<ComprobanteConceptoImpuestosTraslado>();
            ComprobanteConcepto oConcepto = new ComprobanteConcepto();
            ComprobanteConceptoImpuestosTraslado oTraslados = new ComprobanteConceptoImpuestosTraslado();
            ComprobanteConceptoImpuestos oImpuestos = new ComprobanteConceptoImpuestos();
            oConcepto.ClaveProdServ = "84111506";
            oConcepto.NoIdentificacion = numOp;
            oConcepto.Cantidad = 1m;
            oConcepto.ClaveUnidad = "ACT";
            oConcepto.Unidad = "0";
            oConcepto.Descripcion = "DEVOLUCION CON UUID RELACIONADO: " + UUIDRELACIONADO;
            oConcepto.ValorUnitario = Math.Round((totalNegativo / 1.16m), 2);
            oConcepto.Importe = Math.Round((totalNegativo / 1.16m), 2);




            //Aquí pondremos los impuestos
            oTraslados.Base = Math.Round((totalNegativo / 1.16m), 2); //Sin iva
            oTraslados.Impuesto = c_Impuesto.Item002; //Siempre sera esto
            oTraslados.TipoFactor = c_TipoFactor.Tasa;
            oTraslados.TasaOCuota = 0.160000m;
            oTraslados.Importe = Math.Round(totalNegativo - (totalNegativo / 1.16m), 2);

            lstTrasla.Add(oTraslados);
            oImpuestos.Traslados = lstTrasla.ToArray();
            lstTrasla.Clear();

            oConcepto.Impuestos = oImpuestos;

            lstConceptos.Add(oConcepto);
            oComprobante.Conceptos = lstConceptos.ToArray();

            ComprobanteImpuestos oImpuestos2 = new ComprobanteImpuestos();
            ComprobanteImpuestosTraslado oTraslado2 = new ComprobanteImpuestosTraslado();
            List<ComprobanteImpuestosTraslado> lstTraslado2 = new List<ComprobanteImpuestosTraslado>();
            oImpuestos2.TotalImpuestosTrasladados = Math.Round(totalNegativo - (totalNegativo / 1.16m), 2);

            oTraslado2.Impuesto = c_Impuesto.Item002;
            oTraslado2.TipoFactor = c_TipoFactor.Tasa;
            oTraslado2.TasaOCuota = 0.160000m;
            oTraslado2.Importe = Math.Round(totalNegativo - (totalNegativo / 1.16m), 2);
            lstTraslado2.Add(oTraslado2);
            oImpuestos2.Traslados = lstTraslado2.ToArray();
            oComprobante.Impuestos = oImpuestos2;

            //Crear el XML
            XML(oComprobante);

            //xsl //Crear la cadena 
            string cadenaOriginal = "";
            string pathxsl = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\cadenaoriginal_3_3.xslt";
            System.Xml.Xsl.XslCompiledTransform transformador = new System.Xml.Xsl.XslCompiledTransform(true);
            transformador.Load(pathxsl);
            using (StringWriter sw = new StringWriter())
            using (XmlWriter xwo = XmlWriter.Create(sw, transformador.OutputSettings))
            {
                transformador.Transform(pathXML, xwo);
                cadenaOriginal = sw.ToString();
                //MessageBox.Show(cadenaOriginal + "");
            }

            //Sellar el documento
            SelloDigital oselloDigital = new SelloDigital();
            oComprobante.Certificado = oselloDigital.Certificado(pathCer);
            oComprobante.Sello = Test_Mex_Sign_Data(cadenaOriginal);

            //Sobre escribir el xml ya sellado 
            XML(oComprobante);


            //Se instancia el WS de Timbrado.
            ServiceReferenceFac.WSCFDI33Client ServicioTimbrado_FEL = new ServiceReferenceFac.WSCFDI33Client();

            //Se instancia la Respuesta del WS de Timbrado.
            ServiceReferenceFac.RespuestaTFD33 RespuestaTimbrado_FEL = new ServiceReferenceFac.RespuestaTFD33();
            ServiceReferenceFac.RespuestaTFD33 RespuestaServicio_FEL = new ServiceReferenceFac.RespuestaTFD33();
            //Se carga el XML desde archivo.
            XmlDocument DocumentoXML = new XmlDocument();
            //La direccion se sustituira dependiendo de donde se leera el XML.
            DocumentoXML.Load(pathxDinamico + diaCarpeta + numOpPathXMLFac + ".xml");

            //Variable string que contiene el contenido del XML.
            string stringXML = null;
            stringXML = DocumentoXML.OuterXml;
            //Timbrar
            RespuestaTimbrado_FEL = ServicioTimbrado_FEL.TimbrarCFDI(UsuarioFell, ContraseñaFEll, stringXML, numOp);

            //Obteniendo la respuesta se valida que haya sido exitosa.
            if (RespuestaTimbrado_FEL.OperacionExitosa == true)
            {
                MessageBox.Show("ESTADO DE LA NOTA DE CREDITO POR DIFERENCIA: " + RespuestaTimbrado_FEL.Timbre.Estado + System.Environment.NewLine);
                UUIDNUEVO = RespuestaTimbrado_FEL.Timbre.UUID + System.Environment.NewLine;
                observaciones = RespuestaTimbrado_FEL.Timbre.Estado.ToString();
                DocumentoXML.LoadXml(RespuestaTimbrado_FEL.XMLResultado);
                DocumentoXML.Save(pathxDinamico + diaCarpeta + numOp + ".xml");
                //Generar PDF
                RespuestaServicio_FEL = ServicioTimbrado_FEL.ObtenerPDF(UsuarioFell, ContraseñaFEll, UUIDNUEVO, "");
                //Guardo el PDF del CFDi.
                try
                {
                    File.WriteAllBytes(pathxDinamico + diaCarpeta + numOp + ".pdf", Convert.FromBase64String(RespuestaServicio_FEL.PDFResultado));
                    try
                    {
                        //Enviar pdf por correo(Si es que existe) Crear metodo para enviar
                        EnviarCorreo(pathxDinamico + diaCarpeta + numOp + ".pdf", pathxDinamico + diaCarpeta + numOp + ".xml", to);
                    }
                    catch
                    {
                        MessageBox.Show("NO SE ENVIÓ EL CORREO", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception errorPdf)
                {
                    MessageBox.Show("ERROR AL CREAR EL PDF:" + errorPdf, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show("" + RespuestaServicio_FEL.CodigoRespuesta + System.Environment.NewLine +
                    RespuestaServicio_FEL.MensajeError + System.Environment.NewLine +
                     RespuestaServicio_FEL.MensajeErrorDetallado + System.Environment.NewLine);
                }
                facturaOK = 0;
            }
            else
            {
                facturaOK = 1;
                MessageBox.Show("ERROR AL TIMBRAR: " + RespuestaTimbrado_FEL.CodigoRespuesta + System.Environment.NewLine + RespuestaTimbrado_FEL.MensajeError + System.Environment.NewLine + RespuestaTimbrado_FEL.MensajeErrorDetallado + System.Environment.NewLine, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                observaciones = RespuestaTimbrado_FEL.CodigoRespuesta.ToString() + " " + RespuestaTimbrado_FEL.MensajeErrorDetallado.ToString();
            }

        }

        //CANCELACIO POR NEGATIVO
        private void cancelaNegativo()
        {
                
                //Uso CFDI
                if (checkOtro.Checked)
                {

                    usoCFDII = comboOtro.SelectedItem.ToString();
                }
                else
                {
                    if (checkG01.Checked)
                    {
                        usoCFDII = "G01";
                    }
                    if (checkG03.Checked)
                    {
                        usoCFDII = "G03";
                    }
                    if (checkP01.Checked)
                    {
                        usoCFDII = "P01";
                    }
                }
                //Fecha
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "yyyy-MM-dd H:mm:ss";
                fecha = dateTimePicker1.Text;

                
                numOp = encontrarConsecutivo(numOpNC, operaci, "NC");

                //Conexion sql
                String servidor = GetServidor(opServer);
                SqlConnection cn = new SqlConnection(servidor);
                cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                //SqlDataReader lector;

                importe = Convert.ToDouble( totalNegativo);
                Subtotal = (importe / 1.16);
                Total = Convert.ToDouble(totalNegativo);
                totalletras = toText(Total);
                //totalletras
                
                //NOTA DE CREDITO ################################################################

                cmd.CommandText = "INSERT INTO dbo.Facturas([Numero],[Fecha],[Cuenta],[Nombre],[Domicilio],[Interior],[Exterior],[Colonia],[Ciudad],[Estado],[CP],[Correo],[Telefono],[RFC],[Operacion],[Importe],[Descuento],[Porcentaje],[Subtotal],[IVA],[Total],[Letra],[TotalCFDI],[UsoCFDI],[FacturaNoGenerica],[UUID],[Observaciones],[UUIDRelacionado],[UUIDCancelado]) VALUES('" + numOp + "',@FechaHoy,'" + cuenta + "','" + nombre + "',null,null,null,null,null,null,null,'" + email + "','" + telefono + "','" + RFC + "','" + operaci + "','" + importe + "','" + Descuento + "','" + porcentaje + "','" + Subtotal + "','" + IVA + "','" + Total + "','" + totalletras + "',null,'" + usoCFDII + "',null,null,null,null,null);";
                cmd.Parameters.Add("@FechaHoy", SqlDbType.Date).Value = dateTimePicker1.Value;
                cmd.ExecuteNonQuery();

                // TIMBRAR #######################################################################
                

                timbrarNotaCredNegativos();

            if (facturaOK == 0)
            {
                //Actualizar SQL #################################################################
                cmd.CommandText = "update dbo.Facturas set UUID = '" + UUIDNUEVO + "', UUIDRELACIONADO = '" + UUIDRELACIONADO + "', FormaPago ='" + formaPago.ToString() + "', MetodoPago = '" + MetodoPago.ToString() + "' , Observaciones= '" + observaciones + "' where Numero = '" + numOp + "' ";
                cmd.ExecuteNonQuery();
            }
            else
            {
                cmd.CommandText = "delete from dbo.Facturas where Numero = '" + numOp + "'";
            }

            //MessageBox.Show("Se creo la nota de credito de las devolcuiones en esta venta");
            }

        // CANCELACION
        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (UUID == "" || UUID is null)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("NO EXISTE FACTURA CON ESTE NÚMERO DE OPERACIÓN", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                ButtonNota.Enabled = false;
                buttonReFac.Enabled = false;
                timbreCreditoDataGridView.Rows[0].Selected = true;
                DataGridViewSelectedRowCollection row = timbreCreditoDataGridView.SelectedRows;

                if (timbreCreditoDataGridView.CurrentRow == null)
                {
                    //Uso CFDI
                    if (checkOtro.Checked)
                    {

                        usoCFDII = comboOtro.SelectedItem.ToString();
                    }
                    else
                    {
                        if (checkG01.Checked)
                        {
                            usoCFDII = "G01";
                        }
                        if (checkG03.Checked)
                        {
                            usoCFDII = "G03";
                        }
                        if (checkP01.Checked)
                        {
                            usoCFDII = "P01";
                        }
                    }
                    //Fecha
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = "yyyy-MM-dd H:mm:ss";
                    fecha = dateTimePicker1.Text;
                    
                    //fechanew = Convert(dateTimePicker1.Text, DateTime)

                    //Datos
                    operaci = textOperacion.Text;
                    cuenta = textCuenta.Text;
                    RFC = textRFC.Text;
                    nombre = textNombre.Text;
                    email = textEmail.Text;
                    telefono = textTelefono.Text;
                    
                    if (sucursal == "GDL" || sucursal == "ATE" || sucursal == "IXT" || sucursal == "TLA")
                    {
                        numOp = numOp + sucursal + "NC";
                        numOp = numOp.Replace("  ", " ");
                    }
                    else
                    {
                        RFC = "TES030201001"; //RFC TEST
                        //RFC = "CRE840310TC6"; //RFC CRE
                        
                        numOp = encontrarConsecutivo(numOpNC, operaci, "NC");
                        numOp = numOp.Replace("  ", " ");
                    }

                    //MessageBox.Show("NUMERO DE OP1: " + numOp);
                    //Conexion sql
                    String servidor = GetServidor(opServer);
                    SqlConnection cn = new SqlConnection(servidor);
                    cn.Open();
                    SqlCommand cmd = cn.CreateCommand();
                    //SqlDataReader lector;

                    //NOTA DE CREDITO ################################################################
                    if (sucursal == "GDL" || sucursal == "ATE" || sucursal == "IXT" || sucursal == "TLA")
                    {
                        try
                        {
                            cmd.CommandText = "INSERT INTO dbo.Facturas([Numero],[Fecha],[Importe],[Descuento],[UsoCFDI]) VALUES('" + numOp + "',@FechaHoy,'"+importeGlobal+"','"+Descuento+"','"+usoCFDI+"');";
                            cmd.Parameters.Add("@FechaHoy", SqlDbType.Date).Value = Convert.ToDateTime(dateTimePicker1.Value);
                            facturaOK = 0;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception re)
                        {
                            MessageBox.Show("ERROR DE CONEXIÓN: "+re, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            facturaOK = 1;
                        }
                    }
                    else 
                    {
                        try
                        {
                            cmd.CommandText = "INSERT INTO dbo.Facturas([Numero],[Fecha],[Cuenta],[Nombre],[Domicilio],[Interior],[Exterior],[Colonia],[Ciudad],[Estado],[CP],[Correo],[Telefono],[RFC],[Operacion],[Importe],[Descuento],[Porcentaje],[Subtotal],[IVA],[Total],[Letra],[TotalCFDI],[UsoCFDI],[FacturaNoGenerica],[UUID],[Observaciones],[UUIDRelacionado],[UUIDCancelado]) VALUES('" + numOp + "',@FechaHoy,'" + cuenta + "','" + nombre + "',null,null,null,null,null,null,null,'" + email + "','" + telefono + "','" + RFC + "','" + operaci + "','" + importe + "','" + Descuento + "','" + porcentaje + "','" + Subtotal + "','" + IVA + "','" + Total + "','" + totalletras + "',null,'" + usoCFDII + "',null,null,null,null,null);";
                            cmd.Parameters.Add("@FechaHoy", SqlDbType.Date).Value = dateTimePicker1.Value;
                            cmd.ExecuteNonQuery();
                            facturaOK = 0;
                        }
                        catch (Exception re)
                        {
                            MessageBox.Show("ERROR DE CONEXIÓN: " + re, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            facturaOK = 1;
                        }
                        
                    }

                    // TIMBRAR #######################################################################
                    
                    timbrarNotaCred();

                    if (facturaOK == 0)
                    {
                        if (sucursal == "GDL" || sucursal == "ATE" || sucursal == "IXT" || sucursal == "TLA")
                        {
                            //Actualizar SQL #################################################################
                            //MessageBox.Show("NUMERO DE OP2: "+numOp);
                            cmd.CommandText = "update dbo.Facturas set UUID = '" + UUIDNUEVO + "', UUIDRELACIONADO = '" + UUIDRELACIONADO + "', Observaciones= '" + observaciones + "' where Numero = '" + numOp + "' ";
                            cmd.ExecuteNonQuery();
                            if (facturaOK == 0)
                            {
                                //Mover a relacionado
                                cmd.CommandText = "update dbo.Facturas set UUID = null, UUIDRELACIONADO = '" + UUIDRELACIONADO + "' where UUID = '" + UUIDRELACIONADO + "' ";
                                cmd.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            //Actualizar SQL #################################################################
                            cmd.CommandText = "update dbo.Facturas set UUID = '" + UUIDNUEVO + "', UUIDRELACIONADO = '" + UUIDRELACIONADO + "', FormaPago ='" + formaPagoStr + "', MetodoPago = '" + MetodoPago.ToString() + "' , Observaciones= '" + observaciones + "' where Numero = '" + numOp + "' ";
                            cmd.ExecuteNonQuery();

                            if (facturaOK == 0)
                            {
                                //Mover a relacionado
                                cmd.CommandText = "update dbo.Facturas set   UUID = null, UUIDRELACIONADO = '" + UUIDRELACIONADO + "' where Numero = '" + numOpSinCambios + "' ";
                                cmd.ExecuteNonQuery();
                            }
                        }


                        Cursor.Current = Cursors.Default;
                        labelInfo.Text = "NOTA DE CRÉDITO COMPLETA";
                    }
                    else
                    {
                        cmd.CommandText = "delete from dbo.Facturas where Numero = '" + numOp + "'";
                    }
                }
                else
                {
                    Cursor.Current = Cursors.Default;
                    labelInfo.Text = "NO SE PUEDE HACER UNA NOTA POR QUE \nTIENE ABONOS O ENGANCHE YA FATURADOS";

                }
            }
        }

        //ARCHIVO BINARIO
        static bool MakeABinaryFile(string fileName, byte[] data)
        {
            FileStream fs;
            BinaryWriter bw;
            fs = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            bw = new BinaryWriter(fs);
            bw.Write(data);
            bw.Close();
            fs.Close();
            return true;
        }

        public Devolucion()
        {
            InitializeComponent();
            comboOtro.Enabled = false;
        }

        private void facturasBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.facturasBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dbSIADataFact);

        }

        //MAIN LOAD
        private void Devolucion_Load(object sender, EventArgs e)
        {
            
            checkG01.Checked = true;
            comboOtro.Items.Add("G02");
            comboOtro.Items.Add("I02");
            comboOtro.Items.Add("I04");
            comboOtro.Items.Add("I01");
            comboOtro.Items.Add("I03");
            comboOtro.Items.Add("I05");
            comboOtro.Items.Add("I06");
            comboOtro.Items.Add("I07");
            comboOtro.Items.Add("I08");
            comboOtro.Items.Add("D01");
            comboOtro.Items.Add("D02");
            comboOtro.Items.Add("D03");
            comboOtro.Items.Add("D04");
            comboOtro.Items.Add("D05");
            comboOtro.Items.Add("D06");
            comboOtro.Items.Add("D07");
            comboOtro.Items.Add("D08");
            comboOtro.Items.Add("D09");
            comboOtro.Items.Add("D10");

            textNumOp.Enabled = false;
            textCuenta.MaxLength = 10;
            textRFC.MaxLength = 13;
            textNombre.MaxLength = 150;
            textEmail.MaxLength = 50;
            textTelefono.MaxLength = 15;

            
        }

        //NUMERO A LETRAS
        public string enletras(string num)
        {
            string res, dec = "";
            Int64 entero;
            int decimales;
            double nro;

            try

            {
                nro = Convert.ToDouble(num);
            }
            catch
            {
                return "";
            }

            entero = Convert.ToInt64(Math.Truncate(nro));
            decimales = Convert.ToInt32(Math.Round((nro - entero) * 100, 2));
            if (decimales > 0)
            {
                dec = " CON " + decimales.ToString() + "/100";
            }

            res = toText(Convert.ToDouble(entero)) + dec;
            return res;
        }

        //A TEXTO
        private string toText(double value)
        {
            string Num2Text = "";
            value = Math.Truncate(value);
            if (value == 0) Num2Text = "CERO";
            else if (value == 1) Num2Text = "UNO";
            else if (value == 2) Num2Text = "DOS";
            else if (value == 3) Num2Text = "TRES";
            else if (value == 4) Num2Text = "CUATRO";
            else if (value == 5) Num2Text = "CINCO";
            else if (value == 6) Num2Text = "SEIS";
            else if (value == 7) Num2Text = "SIETE";
            else if (value == 8) Num2Text = "OCHO";
            else if (value == 9) Num2Text = "NUEVE";
            else if (value == 10) Num2Text = "DIEZ";
            else if (value == 11) Num2Text = "ONCE";
            else if (value == 12) Num2Text = "DOCE";
            else if (value == 13) Num2Text = "TRECE";
            else if (value == 14) Num2Text = "CATORCE";
            else if (value == 15) Num2Text = "QUINCE";
            else if (value < 20) Num2Text = "DIECI" + toText(value - 10);
            else if (value == 20) Num2Text = "VEINTE";
            else if (value < 30) Num2Text = "VEINTI" + toText(value - 20);
            else if (value == 30) Num2Text = "TREINTA";
            else if (value == 40) Num2Text = "CUARENTA";
            else if (value == 50) Num2Text = "CINCUENTA";
            else if (value == 60) Num2Text = "SESENTA";
            else if (value == 70) Num2Text = "SETENTA";
            else if (value == 80) Num2Text = "OCHENTA";
            else if (value == 90) Num2Text = "NOVENTA";
            else if (value < 100) Num2Text = toText(Math.Truncate(value / 10) * 10) + " Y " + toText(value % 10);
            else if (value == 100) Num2Text = "CIEN";
            else if (value < 200) Num2Text = "CIENTO " + toText(value - 100);
            else if ((value == 200) || (value == 300) || (value == 400) || (value == 600) || (value == 800)) Num2Text = toText(Math.Truncate(value / 100)) + "CIENTOS";
            else if (value == 500) Num2Text = "QUINIENTOS";
            else if (value == 700) Num2Text = "SETECIENTOS";
            else if (value == 900) Num2Text = "NOVECIENTOS";
            else if (value < 1000) Num2Text = toText(Math.Truncate(value / 100) * 100) + " " + toText(value % 100);
            else if (value == 1000) Num2Text = "MIL";
            else if (value < 2000) Num2Text = "MIL " + toText(value % 1000);
            else if (value < 1000000)
            {
                Num2Text = toText(Math.Truncate(value / 1000)) + " MIL";
                if ((value % 1000) > 0) Num2Text = Num2Text + " " + toText(value % 1000);
            }

            else if (value == 1000000) Num2Text = "UN MILLON";
            else if (value < 2000000) Num2Text = "UN MILLON " + toText(value % 1000000);
            else if (value < 1000000000000)
            {
                Num2Text = toText(Math.Truncate(value / 1000000)) + " MILLONES ";
                if ((value - Math.Truncate(value / 1000000) * 1000000) > 0) Num2Text = Num2Text + " " + toText(value - Math.Truncate(value / 1000000) * 1000000);
            }

            else if (value == 1000000000000) Num2Text = "UN BILLON";
            else if (value < 2000000000000) Num2Text = "UN BILLON " + toText(value - Math.Truncate(value / 1000000000000) * 1000000000000);

            else
            {
                Num2Text = toText(Math.Truncate(value / 1000000000000)) + " BILLONES";
                if ((value - Math.Truncate(value / 1000000000000) * 1000000000000) > 0) Num2Text = Num2Text + " " + toText(value - Math.Truncate(value / 1000000000000) * 1000000000000);
            }
            return Num2Text;

        }
        
        //ENCONTRAR Y AGREGAR CONSECUTIVO
        private string encontrarConsecutivo(string varNumero, string varOperacion, string letra)
        {
            //Agregar numero consecutivo
            int sal = 1;
            int contaI = 0;
            if (varNumero.Contains(letra))
            {
                while (sal != 0)
                {
                    if (varNumero.Contains(contaI.ToString()) && contaI >= 100)
                    {
                        contaI++;
                        if (contaI == 1000)
                        {
                            MessageBox.Show("NÚMERO DE NOTAS DE ESTA FACTURA EXCEDIDAS", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            varNumero = varOperacion + letra + contaI.ToString();
                        }
                        sal = 0;
                    }
                    else if (varNumero.Contains("0" + contaI.ToString()) && contaI < 100 && contaI > 9)
                    {
                        contaI++;
                        if (contaI == 100)
                        {
                            varNumero = varOperacion + letra + contaI.ToString();
                        }
                        else
                        {
                            varNumero = varOperacion + letra + "0" + contaI.ToString();
                        }
                        sal = 0;
                    }
                    else if (varNumero.Contains("00" + contaI.ToString()) && contaI <= 9)
                    {
                        contaI++;
                        if (contaI == 10)
                        {
                            varNumero = varOperacion + letra + "0" + contaI.ToString();
                        }
                        else
                        {
                            varNumero = varOperacion + letra + "00" + contaI.ToString();
                        }
                        sal = 0;
                    }
                    contaI++;
                    if (contaI == 999)
                    {
                        sal = 0;
                    }
                }
            }
            else
            {
                varNumero += letra + "001";
            }
            return varNumero;
        }

        //BUSCAR FACTURA
        private void buttonBuscar_Click(object sender, EventArgs e)
        {

            //PROGRAMAR VALIDAR NUMERO DE OPERACION ----------------------------------------------------------- 1.200.1
            buttonReFac.Enabled = true;
            ButtonNota.Enabled = true;
            ButtonFacturar.Enabled = true;
            buttonCorreo.Enabled = true;
            textEmail.Text = null;
            textNombre.Text = null;
            textNumOp.Text = null;
            textCuenta.Text = null;
            textRFC.Text = null;
            textTelefono.Text = null;

            textCuenta.Enabled = false;
            dateTimePicker1.Enabled = false;

            Cursor.Current = Cursors.WaitCursor;
            if (facturasDataGridView.Rows.Count != 0)
            {
                this.facturasDataGridView.DataBindings.Clear();
                facturasDataGridView.Refresh();
            }
            if (det_VentDataGridView.Rows.Count != 0)
            {
                this.det_VentDataGridView.DataBindings.Clear();
                det_VentDataGridView.Refresh();
            }

            texttotal.Text = "TOTAL DE LA VENTA: ";
            labelInfo.Text = "";
            operacion = textOperacion.Text;
            
            
            sucursal = operacion.Substring(4, 2);
            
            
            if (sucursal == "01" || sucursal == "02" || sucursal == "03" || sucursal == "04" || sucursal == "05" || sucursal == "13" || sucursal == "14" || sucursal == "15" || sucursal == "19" || sucursal == "20")
            {
                try
                {
                    //GUADALAJARA
                    this.facturasTableAdapter.Fill(this.dbSIADataFact.Facturas, operacion);
                    this.det_VentTableAdapter.Fill(this.dbSIADataDetalles.Det_Vent, operacion);
                    this.movtosTableAdapter.Fill(this.dbCreditoDataCredito.movtos, operacion);
                    lugarExpedicion = "44280";
                    pathxDinamico = pathxGDL;
                    opServer = 1;
                }
                catch (Exception error)
                {
                }
            }
            else if (sucursal == "06" || sucursal == "07" || sucursal == "08" || sucursal == "09")
            {
                try { 

                //almacen = "02";//ATEQUIZA
                this.facturasTableAdapter.FillByATQ(this.dbSIADataFact.Facturas, operacion);
                this.det_VentTableAdapter.FillByATQ(this.dbSIADataDetalles.Det_Vent, operacion);
                this.movtosTableAdapter.FillByATQ(this.dbCreditoDataCredito.movtos, operacion);
                    lugarExpedicion = "45850";
                    pathxDinamico = pathxATQ;
                    opServer = 2;
                }
                catch (Exception error)
            {
                
            }
        }
            else if (sucursal == "10" || sucursal == "11" || sucursal == "12")
            {
                try { 
                //almacen = "03"; //IXTLAHUACAN
                this.facturasTableAdapter.FillByIXTLA(this.dbSIADataFact.Facturas, operacion);
                this.det_VentTableAdapter.FillByIXTLA(this.dbSIADataDetalles.Det_Vent, operacion);
                this.movtosTableAdapter.FillByIXTLA(this.dbCreditoDataCredito.movtos, operacion);
                    lugarExpedicion = "45860";
                    pathxDinamico = pathxIXT;
                    opServer = 3;
                }
                catch (Exception error)
                {
                    
                }
            }
            else if (sucursal == "21" || sucursal == "22" || sucursal == "23" || sucursal == "24" || sucursal == "25" || sucursal == "26" || sucursal == "27" || sucursal == "28" || sucursal == "29")
            {
                try { 
                //almacen = "04"; //TLAJOMULCO
                this.facturasTableAdapter.FillByTLAJO(this.dbSIADataFact.Facturas, operacion);
                this.det_VentTableAdapter.FillByTLAJO(this.dbSIADataDetalles.Det_Vent, operacion);
                this.movtosTableAdapter.FillByTLAJO(this.dbCreditoDataCredito.movtos, operacion);
                    lugarExpedicion = "45640";
                    pathxDinamico = pathxTLA;
                    opServer = 4;
                }
                catch (Exception error)
                {
                   
                }
            }
            else
            {
                //FACTURA GLOBAL, IDENTIFICAR POR LOS ULTIMOS 3 DIGITOS (12, 13 Y 14) ATE, GDL, IXT, TLA
                sucursal = operacion.Substring(11, 3);
                if(sucursal == "GDL")
                {
                    //Llenar gdl
                    this.facturasTableAdapter.Fill(this.dbSIADataFact.Facturas, operacion);
                    lugarExpedicion = "44280";
                    pathxDinamico = pathxGDL;
                    opServer = 1;
                }
                else if(sucursal == "ATE")
                {
                    //Llenar ate
                    this.facturasTableAdapter.FillByATQ(this.dbSIADataFact.Facturas, operacion);
                    lugarExpedicion = "45850";
                    pathxDinamico = pathxATQ;
                    opServer = 2;
                }
                else if(sucursal == "IXT")
                {
                    //Llenar ixtla
                    this.facturasTableAdapter.FillByIXTLA(this.dbSIADataFact.Facturas, operacion);
                    lugarExpedicion = "45860";
                    pathxDinamico = pathxIXT;
                    opServer = 3;
                }
                else if(sucursal == "TLA")
                {
                    //Llenar tlajo
                    this.facturasTableAdapter.FillByTLAJO(this.dbSIADataFact.Facturas, operacion);
                    lugarExpedicion = "45640";
                    pathxDinamico = pathxTLA;
                    opServer = 4;
                }

            }

            //Crear carpeta si no existe 
            DateTime dateTime = DateTime.Now.Date;
            string fechaC = dateTime.ToString("dd-MM-yyyy");
            diaCarpeta = fechaC;
            //MessageBox.Show("FECHA: " + diaCarpeta);
            if (Directory.Exists(pathxDinamico + diaCarpeta))
            {
               // MessageBox.Show("La carpeta ya existe : " + pathxDinamico + diaCarpeta);
            }
            else
            {
                try
                {

                    Directory.CreateDirectory(pathxDinamico + diaCarpeta);
                    //MessageBox.Show("La carpeta se ha creado: " + pathxDinamico + diaCarpeta); 
                }
                catch (Exception error)
                {
                    MessageBox.Show("ERROR DE CONEXIÓN: " + error, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            diaCarpeta += @"\";

            //TERMINA IDENTIFICAR

            //PINTAR 
            foreach (DataGridViewRow vFila in facturasDataGridView.Rows)
            {
                vFila.Cells[30].Style.Format = "n2";
            }

                foreach (DataGridViewRow vFila in det_VentDataGridView.Rows)
            {
                vFila.Cells[13].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                vFila.Cells[14].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                vFila.Cells[15].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                vFila.Cells[17].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                vFila.Cells[18].Style.Format = "n2";

                int tempo = Int32.Parse(vFila.Cells[14].Value.ToString());
                int tempo2 = Int32.Parse(vFila.Cells[15].Value.ToString());

                if (tempo < 0)
                {
                    vFila.Cells[14].Style.BackColor = Color.IndianRed;
                }
                else if (tempo2 < 0)
                {
                    vFila.Cells[15].Style.BackColor = Color.IndianRed;
                }

            }
            //TERMINA PINTAR
            
            try
            {
                facturasDataGridView.Rows[0].Selected = true;
                DataGridViewSelectedRowCollection row = facturasDataGridView.SelectedRows;
                //UUID ANTERIOR
                UUIDRELACIONADO = row[0].Cells[25].Value.ToString();
                UUID = UUIDRELACIONADO;
                //MessageBox.Show("UUID: " + UUID);
                importeGlobal = row[0].Cells[17].Value.ToString();
                totald = Convert.ToDecimal(row[0].Cells[17].Value);
                usoCFDI = row[0].Cells[23].Value.ToString();
                texttotal.Text = "TOTAL DE LA VENTA: " + row[0].Cells[17].Value.ToString();
                numOp = row[0].Cells[2].Value.ToString();
                numOpSinCambios = row[0].Cells[2].Value.ToString();
                numOpProd = numOp;
                numOpPathXMLFac = numOp;
                pathXML = pathxDinamico + diaCarpeta + numOpPathXMLFac + ".xml";
                formaPagoStr = row[0].Cells[31].Value.ToString();
                auxformapago = Convert.ToDouble(row[0].Cells[31].Value);
                formaPago = Convert.ToUInt16(auxformapago);
                MetodoPago = row[0].Cells[32].Value.ToString();
                Total = Convert.ToDouble(row[0].Cells[30].Value);
                to = row[0].Cells[14].Value.ToString();
                RFCAnterior = row[0].Cells[16].Value.ToString();
                importe = Convert.ToDouble(row[0].Cells[17].Value);
                Descuento = Convert.ToDouble(row[0].Cells[18].Value);
                descuentod = Convert.ToDecimal(Descuento);
                porcentaje = Convert.ToDouble(row[0].Cells[19].Value);
                Subtotal = Convert.ToDouble(row[0].Cells[20].Value);
                subtotald = Convert.ToDecimal(Subtotal);
                IVA = Convert.ToDouble(row[0].Cells[21].Value);
                texttotal.Text = "TOTAL DE LA VENTA: " + row[0].Cells[30].Value.ToString();
                textCuenta.Text = row[0].Cells[6].Value.ToString();
                textNombre.Text = row[0].Cells[3].Value.ToString();
                textNumOp.Text = row[0].Cells[2].Value.ToString();
                textEmail.Text = row[0].Cells[14].Value.ToString();
                textTelefono.Text = row[0].Cells[15].Value.ToString();
                
                //A letras
                string totalnum = row[0].Cells[30].Value.ToString();
                totalletras = enletras(totalnum);
                totalletras += " PESOS 00/100 M.N.";
                totald = Convert.ToDecimal(Total);
            }
            catch
            {
                if (sucursal == "GDL" || sucursal == "ATE" || sucursal == "IXT" || sucursal == "TLA")
                {
                    Cursor.Current = Cursors.Default;
                    labelInfo.Text = "ESTA ES FACTURA GLOBAL";
                    texttotal.Text = "TOTAL DE LA FACTURA: "+importeGlobal;
                    
                    buttonReFac.Enabled = false;
                    ButtonNota.Enabled = true;
                }
                else
                {
                    Cursor.Current = Cursors.Default;
                    labelInfo.Text = "NO EXISTE FACTURA GLOBAL CON ESE NUMERO \nDE OPERACION, POR FAVOR PROPORCIONE UN \nNUMERO DE OPERACION CORRECTO";
                    ButtonFacturar.Enabled = false;
                    buttonReFac.Enabled = false;
                    ButtonNota.Enabled = false;
                }

            }
            try
            {
                int salir = 0;
                int conta1=1;
                while (facturasDataGridView.Rows[facturasDataGridView.Rows.Count - conta1].ToString() != null && salir == 0) {
                    facturasDataGridView.Rows[facturasDataGridView.Rows.Count - conta1].Selected = true;
                    DataGridViewSelectedRowCollection row = facturasDataGridView.SelectedRows;
                    numOpNC = row[0].Cells[4].Value.ToString();
                    conta1++;
                    
                    if (numOpNC.Contains("NC"))
                    {
                        salir = 1;
                        noRefactura = 1;
                    }
                    else
                    {
                        numOpFac = numOp;
                    }
                }
            }
            catch
            {
            }
            try
            {
                int salir = 0;
                int conta1 = 1;
                while (facturasDataGridView.Rows[facturasDataGridView.Rows.Count - conta1].ToString() != null && salir == 0)
                {
                    facturasDataGridView.Rows[facturasDataGridView.Rows.Count - conta1].Selected = true;
                    DataGridViewSelectedRowCollection row = facturasDataGridView.SelectedRows;
                    numOpFac = row[0].Cells[4].Value.ToString();
                    conta1++;

                    if (numOpFac.Contains("F"))
                    {
                        salir = 1;
                        noRefactura = 1;
                    }
                    else
                    {
                        numOpFac = numOp;
                    }

                    if (numOpFac.Contains("FC"))
                    {
                        noRefactura = 1;
                    }
                }
            }
            catch
            {

            }
            Cursor.Current = Cursors.Default;

            if (noRefactura == 1)
            {
                ButtonFacturar.Enabled = false;
                buttonReFac.Enabled = false;
                ButtonNota.Enabled = false;

            }
            if(RFCAnterior != "XAXX010101000")
            {
                ButtonFacturar.Enabled = false;
                buttonReFac.Enabled = false;
                ButtonNota.Enabled = false;
                labelInfo.Text = "ESTA OPERACIÓN TIENE RFC NOMINATIVO";
            }
        }

        //REFACTURAR
        private void buttonReFac_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (UUID == "" || UUID is null)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("NO EXISTE FACTURA CON ESTE NÚMERO DE OPERACIÓN", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                ButtonNota.Enabled = false;
                buttonReFac.Enabled = false;
                timbreCreditoDataGridView.Rows[0].Selected = true;
                DataGridViewSelectedRowCollection row = timbreCreditoDataGridView.SelectedRows;
                tieneAbono = false;
                //1.300.1
                if (timbreCreditoDataGridView.CurrentRow != null) //Revisar abonos para timbrar despues de facturar
                {
                    tieneAbono = true;
                }

                //Uso CFDI
                if (checkOtro.Checked)
                {

                    usoCFDII = comboOtro.SelectedItem.ToString();
                }
                else
                {
                    if (checkG01.Checked)
                    {
                        usoCFDII = "G01";
                    }
                    if (checkG03.Checked)
                    {
                        usoCFDII = "G03";
                    }
                    if (checkP01.Checked)
                    {
                        usoCFDII = "P01";
                    }
                }
                //Fecha
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "yyyy-MM-dd H:mm:ss";
                fecha = dateTimePicker1.Text;

                //Datos
                operaci = textOperacion.Text;
                cuenta = textCuenta.Text;
                RFC = textRFC.Text;
                nombre = textNombre.Text;
                email = textEmail.Text;
                telefono = textTelefono.Text;

                numOp = encontrarConsecutivo(numOpNC, operaci, "NC");
                numOp = numOp.Replace("  ", " ");
                //Conexion sql
                String servidor = GetServidor(opServer);
                SqlConnection cn = new SqlConnection(servidor);
                cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                //SqlDataReader lector;

                //NOTA DE CREDITO ################################################################

                cmd.CommandText = "INSERT INTO dbo.Facturas([Numero],[Fecha],[Cuenta],[Nombre],[Domicilio],[Interior],[Exterior],[Colonia],[Ciudad],[Estado],[CP],[Correo],[Telefono],[RFC],[Operacion],[Importe],[Descuento],[Porcentaje],[Subtotal],[IVA],[Total],[Letra],[TotalCFDI],[UsoCFDI],[FacturaNoGenerica],[UUID],[Observaciones],[UUIDRelacionado],[UUIDCancelado]) VALUES('" + numOp + "',@FechaHoy,'" + cuenta + "','" + nombre + "',null,null,null,null,null,null,null,'" + email + "','" + telefono + "','" + RFCAnterior + "','" + operaci + "','" + importe + "','" + Descuento + "','" + porcentaje + "','" + Subtotal + "','" + IVA + "','" + Total + "','" + totalletras + "',null,'" + usoCFDII + "',null,null,null,null,null);";
                cmd.Parameters.Add("@FechaHoy", SqlDbType.Date).Value = dateTimePicker1.Value;
                cmd.ExecuteNonQuery();

                // TIMBRAR #######################################################################
                timbrarNotaCred();

                //Actualizar SQL #################################################################
                facturaOK = 1;
                if (facturaOK == 0)
                {
                    cmd.CommandText = "update dbo.Facturas set UUID = '" + UUIDNUEVO + "', UUIDRELACIONADO = '" + UUIDRELACIONADO + "', FormaPago ='" + formaPagoStr + "', MetodoPago = '" + MetodoPago.ToString() + "' , Observaciones= '" + observaciones + "' where Numero = '" + numOp + "' ";
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    //En caso de que no se haga la factura se borra el registro
                    cmd.CommandText = "delete from dbo.Facturas where Numero = '" + numOp + "'";
                }

                //####################################-----------------------------------------

                numOp = encontrarConsecutivo(numOpFac, operaci, "F");
                numOp = numOp.Replace("  ", " ");
                //FACTURA AL SQL ################################################################# NO tiene uuid relacionado, solo va el nuevo
                cmd.CommandText = "INSERT INTO dbo.Facturas([Numero],[Fecha],[Cuenta],[Nombre],[Domicilio],[Interior],[Exterior],[Colonia],[Ciudad],[Estado],[CP],[Correo],[Telefono],[RFC],[Operacion],[Importe],[Descuento],[Porcentaje],[Subtotal],[IVA],[Total],[Letra],[TotalCFDI],[UsoCFDI],[FacturaNoGenerica],[UUID],[Observaciones],[UUIDRelacionado],[UUIDCancelado]) VALUES('" + numOp + "',@Fech,'" + cuenta + "','" + nombre + "',null,null,null,null,null,null,null,'" + email + "','" + telefono + "','" + RFC + "','" + operaci + "','" + importe + "','" + Descuento + "','" + porcentaje + "','" + Subtotal + "','" + IVA + "','" + Total + "','" + totalletras + "',null,'" + usoCFDII + "',null,null,null,null,null);";
                cmd.Parameters.Add("@Fech", SqlDbType.Date).Value = dateTimePicker1.Value;
                cmd.ExecuteNonQuery();

                //TIMBRAR ########################################################################
                timbrarFactura();//Enviar correo al terminar factura


                if (facturaOK == 0)
                {
                    //Actualizar SQL #################################################################
                    cmd.CommandText = "update dbo.Facturas set UUID = '" + UUIDNUEVO + "', FormaPago ='" + formaPagoStr + "', MetodoPago = '" + MetodoPago.ToString() + "', Observaciones= '" + observaciones + "' where Numero = '" + numOp + "' ";
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    //En caso de que no se haga la factura se borra el registro
                    cmd.CommandText = "delete from dbo.Facturas where Numero = '" + numOp + "'";
                }


                if (facturaOK == 0) {
                    //Mover a relacionado#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
                    //1.200.1
                    cmd.CommandText = "update dbo.Facturas set UUIDCANCELADO= '" + UUIDRELACIONADO + "', Numero = '" + numOpSinCambios + "FC" + "' where Numero = '" + numOpSinCambios + "' ";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "update dbo.Facturas set Numero = '" + numOpSinCambios + "' where Numero = '" + numOp + "' ";
                    cmd.ExecuteNonQuery();
                }

                if (totalNegativo > 0)
                {
                    cancelaNegativo();//Enviar correo al terminar cancelación por negativos

                }
                //1.300.1
                if (tieneAbono == true)
                {
                    //TIMBRAR ABONOS
                }
                labelInfo.Text = "RE FACTURACIÓN COMPLETA";
            }
            Cursor.Current = Cursors.Default;
        }

        //TIMBRAR ABONOS Y ENGANCHE 1.300.1
        private void TimbrarPagos()
        {

        }

        //OBTENER CERTIFICADO
        private string ObtenerCertificado()
        {
            //Obtener el numero
            string numeroCertificado="", aa, b, c;
            
            try
            {
                XSDToXML.Utils.SelloDigital.leerCER(pathCer, out aa, out b, out c, out numeroCertificado);
            }
            catch
            {

            }
            return numeroCertificado;
        }

        //NOTA DE CREDITO
        private void timbrarNotaCred()
        {
            try
            {
                facturaOK = 0;
                //Crear XML de nota de credito
                Comprobante oComprobante = new Comprobante();

                oComprobante.Version = "3.3";
                oComprobante.Folio = numOp;//Numero de operacion 

                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "yyyy-MM-ddTHH:mm:ss";
                oComprobante.Fecha = Convert.ToDateTime(dateTimePicker1.Text.ToString()); //Formato que debe llevar 
                                                                                          // CASE PARA FORMA DE PAGO

                switch (formaPago)
                {
                    case (99):
                        oComprobante.FormaPago = c_FormaPago.Item99;
                        break;
                    case (01):
                        oComprobante.FormaPago = c_FormaPago.Item01;
                        break;
                    case (04):
                        oComprobante.FormaPago = c_FormaPago.Item04;
                        break;
                    case (28):
                        oComprobante.FormaPago = c_FormaPago.Item28;
                        break;
                }
                oComprobante.NoCertificado = ObtenerCertificado();
                oComprobante.SubTotal = Math.Round((totald / 1.16m), 2);
                if (descuentod == 0)
                {
                }
                else
                {
                    oComprobante.Descuento = descuentod;
                }
                oComprobante.Moneda = c_Moneda.MXN;
                oComprobante.Total = Math.Round(totald, 2);
                oComprobante.TipoDeComprobante = c_TipoDeComprobante.E; //Siempre egreso en nota de credito
                if (MetodoPago == "PUE")
                {
                    oComprobante.MetodoPago = c_MetodoPago.PUE;
                }
                else
                {
                    oComprobante.MetodoPago = c_MetodoPago.PPD;
                }
                oComprobante.LugarExpedicion = lugarExpedicion;

                ComprobanteCfdiRelacionados oComprobanteCFDI = new ComprobanteCfdiRelacionados();

                oComprobanteCFDI.TipoRelacion = c_TipoRelacion.Item02; // Siempre va este

                ComprobanteCfdiRelacionadosCfdiRelacionado oComprobanteCFDIRel = new ComprobanteCfdiRelacionadosCfdiRelacionado();
                List<ComprobanteCfdiRelacionadosCfdiRelacionado> lstCfdiRelacRelac = new List<ComprobanteCfdiRelacionadosCfdiRelacionado>();

                oComprobanteCFDIRel.UUID = UUIDRELACIONADO;
                lstCfdiRelacRelac.Add(oComprobanteCFDIRel);
                oComprobanteCFDI.CfdiRelacionado = lstCfdiRelacRelac.ToArray();
                oComprobante.CfdiRelacionados = oComprobanteCFDI;


                ComprobanteEmisor oEmisor = new ComprobanteEmisor();
                oEmisor.RegimenFiscal = c_RegimenFiscal.Item601; //Siempre va este
                ComprobanteReceptor oReceptor = new ComprobanteReceptor();
                
                //PRUEBAS
                oEmisor.Rfc = "TES030201001";
                oEmisor.Nombre = "temporal";
                
                //PRUEBAS
                oReceptor.Nombre = "temporal";
                oReceptor.Rfc = "TEST010203001";
                /*
                //REAL
                oEmisor.Rfc = "CRE840310TC6";
                oEmisor.Nombre = "Comercial Del Retiro SA de CV";

                //REAL
                oReceptor.Nombre = "PUBLICO EN GENERAL";
                oReceptor.Rfc = RFCAnterior;

                /*
                if (RFCAnterior == "XAXX010101000")
                {
                    //REAL
                    oReceptor.Nombre = "PUBLICO EN GENERAL";
                    oReceptor.Rfc = RFCAnterior;
                }
                else
                {
                    //REAL
                    oReceptor.Nombre = nombre;
                    oReceptor.Rfc = RFCAnterior;
                }
                */
                

                //Asignar emisor y receptor
                oComprobante.Emisor = oEmisor;
                oComprobante.Receptor = oReceptor;

                //Switch uso CFDI
                switch (usoCFDI)
                {
                    case "P01":
                        oReceptor.UsoCFDI = c_UsoCFDI.P01;
                        break;
                    case "G01":
                        oReceptor.UsoCFDI = c_UsoCFDI.G01;
                        break;
                    case "G02":
                        oReceptor.UsoCFDI = c_UsoCFDI.G02;
                        break;
                    case "G03":
                        oReceptor.UsoCFDI = c_UsoCFDI.G03;
                        break;
                    case "I01":
                        oReceptor.UsoCFDI = c_UsoCFDI.I01;
                        break;
                    case "I02":
                        oReceptor.UsoCFDI = c_UsoCFDI.I02;
                        break;
                    case "I03":
                        oReceptor.UsoCFDI = c_UsoCFDI.I03;
                        break;
                    case "I04":
                        oReceptor.UsoCFDI = c_UsoCFDI.I04;
                        break;
                    case "I05":
                        oReceptor.UsoCFDI = c_UsoCFDI.I05;
                        break;
                    case "I06":
                        oReceptor.UsoCFDI = c_UsoCFDI.I06;
                        break;
                    case "I07":
                        oReceptor.UsoCFDI = c_UsoCFDI.I07;
                        break;
                    case "I08":
                        oReceptor.UsoCFDI = c_UsoCFDI.I08;
                        break;
                    case "D01":
                        oReceptor.UsoCFDI = c_UsoCFDI.D01;
                        break;
                    case "D02":
                        oReceptor.UsoCFDI = c_UsoCFDI.D02;
                        break;
                    case "D03":
                        oReceptor.UsoCFDI = c_UsoCFDI.D03;
                        break;
                    case "D04":
                        oReceptor.UsoCFDI = c_UsoCFDI.D04;
                        break;
                    case "D05":
                        oReceptor.UsoCFDI = c_UsoCFDI.D05;
                        break;
                    case "D06":
                        oReceptor.UsoCFDI = c_UsoCFDI.D06;
                        break;
                    case "D07":
                        oReceptor.UsoCFDI = c_UsoCFDI.D07;
                        break;
                    case "D08":
                        oReceptor.UsoCFDI = c_UsoCFDI.D08;
                        break;
                    case "D09":
                        oReceptor.UsoCFDI = c_UsoCFDI.D09;
                        break;
                    case "D10":
                        oReceptor.UsoCFDI = c_UsoCFDI.D10;
                        break;

                }

                decimal totaldd = totald - (totald * 0.16m);

                List<ComprobanteConcepto> lstConceptos = new List<ComprobanteConcepto>();
                List<ComprobanteConceptoImpuestos> lstimpuesto = new List<ComprobanteConceptoImpuestos>();
                List<ComprobanteConceptoImpuestosTraslado> lstTrasla = new List<ComprobanteConceptoImpuestosTraslado>();
                ComprobanteConcepto oConcepto = new ComprobanteConcepto();
                ComprobanteConceptoImpuestosTraslado oTraslados = new ComprobanteConceptoImpuestosTraslado();
                ComprobanteConceptoImpuestos oImpuestos = new ComprobanteConceptoImpuestos();
                oConcepto.ClaveProdServ = "84111506";
                oConcepto.NoIdentificacion = numOp;
                oConcepto.Cantidad = 1m;
                oConcepto.ClaveUnidad = "ACT";
                oConcepto.Unidad = "0";
                oConcepto.Descripcion = "REFACTURACIÓN CON UUID RELACIONADO: " + UUIDRELACIONADO;
                oConcepto.ValorUnitario = Math.Round((totald / 1.16m), 2);
                oConcepto.Importe = Math.Round((totald / 1.16m), 2);


                //Aquí pondremos los impuestos
                oTraslados.Base = Math.Round((totald / 1.16m), 2); //Sin iva
                oTraslados.Impuesto = c_Impuesto.Item002; //Siempre sera esto
                oTraslados.TipoFactor = c_TipoFactor.Tasa;
                oTraslados.TasaOCuota = 0.160000m;
                oTraslados.Importe = Math.Round(totald - (totald / 1.16m), 2);

                lstTrasla.Add(oTraslados);
                oImpuestos.Traslados = lstTrasla.ToArray();
                lstTrasla.Clear();

                oConcepto.Impuestos = oImpuestos;

                lstConceptos.Add(oConcepto);
                oComprobante.Conceptos = lstConceptos.ToArray();

                ComprobanteImpuestos oImpuestos2 = new ComprobanteImpuestos();
                ComprobanteImpuestosTraslado oTraslado2 = new ComprobanteImpuestosTraslado();
                List<ComprobanteImpuestosTraslado> lstTraslado2 = new List<ComprobanteImpuestosTraslado>();
                oImpuestos2.TotalImpuestosTrasladados = Math.Round(totald - (totald / 1.16m), 2);

                oTraslado2.Impuesto = c_Impuesto.Item002;
                oTraslado2.TipoFactor = c_TipoFactor.Tasa;
                oTraslado2.TasaOCuota = 0.160000m;
                oTraslado2.Importe = Math.Round(totald - (totald / 1.16m), 2);
                lstTraslado2.Add(oTraslado2);
                oImpuestos2.Traslados = lstTraslado2.ToArray();
                oComprobante.Impuestos = oImpuestos2;

                //Crear el XML 
                XML(oComprobante);
                //xsl //Crear la cadena 
                string cadenaOriginal = "";
                string pathxsl = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\cadenaoriginal_3_3.xslt";
                System.Xml.Xsl.XslCompiledTransform transformador = new System.Xml.Xsl.XslCompiledTransform(true);
                transformador.Load(pathxsl);
                using (StringWriter sw = new StringWriter())
                using (XmlWriter xwo = XmlWriter.Create(sw, transformador.OutputSettings))
                {
                    transformador.Transform(pathXML, xwo);
                    cadenaOriginal = sw.ToString();
                    //MessageBox.Show(cadenaOriginal + "");
                }

                //Sellar el documento
                SelloDigital oselloDigital = new SelloDigital();
                oComprobante.Certificado = oselloDigital.Certificado(pathCer);
                oComprobante.Sello = Test_Mex_Sign_Data(cadenaOriginal);

                //Sobre escribir el xml ya sellado 
                XML(oComprobante);


                //Se instancia el WS de Timbrado.
                ServiceReferenceFac.WSCFDI33Client ServicioTimbrado_FEL = new ServiceReferenceFac.WSCFDI33Client();

                //Se instancia la Respuesta del WS de Timbrado.
                ServiceReferenceFac.RespuestaTFD33 RespuestaTimbrado_FEL = new ServiceReferenceFac.RespuestaTFD33();

                //Se carga el XML desde archivo.
                XmlDocument DocumentoXML = new XmlDocument();
                //La direccion se sustituira dependiendo de donde se leera el XML.

                DocumentoXML.Load(pathxDinamico + diaCarpeta + numOpPathXMLFac + ".xml");

                //Variable string que contiene el contenido del XML.
                string stringXML = null;
                stringXML = DocumentoXML.OuterXml;
                //Timbrar
                RespuestaTimbrado_FEL = ServicioTimbrado_FEL.TimbrarCFDI(UsuarioFell, ContraseñaFEll, stringXML, numOp);

                //Obteniendo la respuesta se valida que haya sido exitosa.
                if (RespuestaTimbrado_FEL.OperacionExitosa == true)
                {
                    MessageBox.Show("ESTADO DE LA NOTA DE CREDITO: " + RespuestaTimbrado_FEL.Timbre.Estado + System.Environment.NewLine, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    UUIDNUEVO = RespuestaTimbrado_FEL.Timbre.UUID + System.Environment.NewLine;
                    observaciones = RespuestaTimbrado_FEL.Timbre.Estado.ToString();
                    DocumentoXML.LoadXml(RespuestaTimbrado_FEL.XMLResultado);
                    DocumentoXML.Save(pathxDinamico + diaCarpeta + numOp + ".xml");
                    facturaOK = 0;
                }
                else
                {
                    MessageBox.Show("ERROR: " + RespuestaTimbrado_FEL.CodigoRespuesta + System.Environment.NewLine + RespuestaTimbrado_FEL.MensajeError + System.Environment.NewLine + RespuestaTimbrado_FEL.MensajeErrorDetallado + System.Environment.NewLine, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    observaciones = RespuestaTimbrado_FEL.CodigoRespuesta.ToString() + " " + RespuestaTimbrado_FEL.MensajeErrorDetallado.ToString();
                    facturaOK = 1;
                }
            }
            catch (Exception error)
            {
                facturaOK = 1;
                MessageBox.Show("ERROR DE SISTEMA, NO SE PUDO CREAR LA NOTA DE CREDITO" + Environment.NewLine + error, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //FACTURAR
        private void timbrarFactura()
        {
            try
            {
                decimal totaltotal = 0;
                Comprobante oComprobante = new Comprobante();
                oComprobante.Version = "3.3";
                oComprobante.Folio = numOp; //Numero de operacion
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "yyyy-MM-ddTHH:mm:ss";
                oComprobante.Fecha = Convert.ToDateTime(dateTimePicker1.Text.ToString()); //Formato que debe llevar 
                                                                                          //CASE PARA FORMA DE PAGO

                switch (formaPago)
                {
                    case (99):
                        oComprobante.FormaPago = c_FormaPago.Item99;
                        break;
                    case (01):
                        oComprobante.FormaPago = c_FormaPago.Item01;
                        break;
                    case (04):
                        oComprobante.FormaPago = c_FormaPago.Item04;
                        break;
                    case (28):
                        oComprobante.FormaPago = c_FormaPago.Item28;
                        break;
                    case (02):
                        oComprobante.FormaPago = c_FormaPago.Item02;
                        break;
                    case (03):
                        oComprobante.FormaPago = c_FormaPago.Item03;
                        break;
                }

                oComprobante.NoCertificado = ObtenerCertificado();

                if (descuentod == 0)
                {
                }
                else
                {
                    oComprobante.Descuento = descuentod;
                }
                oComprobante.Moneda = c_Moneda.MXN;

                oComprobante.TipoDeComprobante = c_TipoDeComprobante.I; //Siempre ingreso en factura

                if (MetodoPago == "PUE")
                {
                    oComprobante.MetodoPago = c_MetodoPago.PUE;
                }
                else
                {
                    oComprobante.MetodoPago = c_MetodoPago.PPD;
                }

                //oComprobante.CondicionesDePago = "Pago"; //Condiciones de pago. No es obligatorio

                oComprobante.LugarExpedicion = lugarExpedicion;

                ComprobanteEmisor oEmisor = new ComprobanteEmisor();
                oEmisor.RegimenFiscal = c_RegimenFiscal.Item601; //Siempre el mismo
                ComprobanteReceptor oReceptor = new ComprobanteReceptor();
                
                /*
                //REALES
                oEmisor.Rfc = "CRE840310TC6";
                oEmisor.Nombre = "COMERCIAL DEL RETIRO SA DE CV";

                //REALES
                oReceptor.Nombre = nombre;
                oReceptor.Rfc = RFC;
                */
                
                
                //PRUEBA
                oEmisor.Rfc = "TES030201001";
                oEmisor.Nombre = "temporal";

                //PRUEBA
                oReceptor.Nombre = "temporal";
                oReceptor.Rfc = "TEST010203001";
                     

                oComprobante.Emisor = oEmisor;
                oComprobante.Receptor = oReceptor;

                //Switch uso CFDI
                switch (usoCFDII)
                {
                    case "P01":
                        oReceptor.UsoCFDI = c_UsoCFDI.P01;
                        break;
                    case "G01":
                        oReceptor.UsoCFDI = c_UsoCFDI.G01;
                        break;
                    case "G02":
                        oReceptor.UsoCFDI = c_UsoCFDI.G02;
                        break;
                    case "G03":
                        oReceptor.UsoCFDI = c_UsoCFDI.G03;
                        break;
                    case "I01":
                        oReceptor.UsoCFDI = c_UsoCFDI.I01;
                        break;
                    case "I02":
                        oReceptor.UsoCFDI = c_UsoCFDI.I02;
                        break;
                    case "I03":
                        oReceptor.UsoCFDI = c_UsoCFDI.I03;
                        break;
                    case "I04":
                        oReceptor.UsoCFDI = c_UsoCFDI.I04;
                        break;
                    case "I05":
                        oReceptor.UsoCFDI = c_UsoCFDI.I05;
                        break;
                    case "I06":
                        oReceptor.UsoCFDI = c_UsoCFDI.I06;
                        break;
                    case "I07":
                        oReceptor.UsoCFDI = c_UsoCFDI.I07;
                        break;
                    case "I08":
                        oReceptor.UsoCFDI = c_UsoCFDI.I08;
                        break;
                    case "D01":
                        oReceptor.UsoCFDI = c_UsoCFDI.D01;
                        break;
                    case "D02":
                        oReceptor.UsoCFDI = c_UsoCFDI.D02;
                        break;
                    case "D03":
                        oReceptor.UsoCFDI = c_UsoCFDI.D03;
                        break;
                    case "D04":
                        oReceptor.UsoCFDI = c_UsoCFDI.D04;
                        break;
                    case "D05":
                        oReceptor.UsoCFDI = c_UsoCFDI.D05;
                        break;
                    case "D06":
                        oReceptor.UsoCFDI = c_UsoCFDI.D06;
                        break;
                    case "D07":
                        oReceptor.UsoCFDI = c_UsoCFDI.D07;
                        break;
                    case "D08":
                        oReceptor.UsoCFDI = c_UsoCFDI.D08;
                        break;
                    case "D09":
                        oReceptor.UsoCFDI = c_UsoCFDI.D09;
                        break;
                    case "D10":
                        oReceptor.UsoCFDI = c_UsoCFDI.D10;
                        break;

                }

                List<ComprobanteConcepto> lstConceptos = new List<ComprobanteConcepto>();
                List<ComprobanteConceptoImpuestos> lstimpuesto = new List<ComprobanteConceptoImpuestos>();
                List<ComprobanteConceptoImpuestosTraslado> lstTrasla = new List<ComprobanteConceptoImpuestosTraslado>();

                int conta1 = 1;
                string auxImporte;
                string codigoArticulo, descripcion;
                decimal cantidad, valorUni, descuento, importe;

                try
                {
                    while (det_VentDataGridView.Rows[det_VentDataGridView.Rows.Count - conta1].ToString() != null)
                    {
                        ComprobanteConcepto oConcepto = new ComprobanteConcepto();

                        //Seleccionar tambien el primero 
                        det_VentDataGridView.Rows[det_VentDataGridView.Rows.Count - conta1].Selected = false;
                        det_VentDataGridView.Rows[det_VentDataGridView.Rows.Count - conta1].Selected = true;



                        DataGridViewSelectedRowCollection row = det_VentDataGridView.SelectedRows;
                        if (Convert.ToInt16(row[0].Cells[16].Value) > 0)
                        {
                            codigoArticulo = row[0].Cells[1].Value.ToString();
                            descripcion = row[0].Cells[2].Value.ToString();
                            cantidad = Convert.ToDecimal(row[0].Cells[13].Value);
                            valorUni = Convert.ToDecimal(row[0].Cells[7].Value);
                            descuento = Convert.ToDecimal(row[0].Cells[9].Value);
                            importe = Convert.ToDecimal(row[0].Cells[17].Value);
                            this.articulosTableAdapter.Fill(this.dbSIADataSetArt.Articulos, codigoArticulo);
                            foreach (DataRow ro in this.dbSIADataSetArt.Articulos)
                            {
                                if (ro[3] != null)
                                {
                                    claveProdServ = (ro[3].ToString());
                                }
                                else
                                {
                                    MessageBox.Show("ESTE ARTÍCULO NO TIENE CLAVE DE PRODUCTOS Y SERVICIOS, CONTACTE A SISTEMAS", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    MessageBox.Show("EL PROGRAMA SE CERRARÁ, PROPORCIONE ESTÁ CLAVE DE PRODUCTOS Y SERVICIOS A SISTEMAS: "+ claveProdServ, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    claveProdServ = null;
                                    facturaOK = 1;
                                }

                            }

                            oConcepto.ClaveProdServ = claveProdServ;

                            oConcepto.Cantidad = Math.Round(cantidad, 2);
                            oConcepto.ClaveUnidad = "H87"; //Codigo para especificar que es por pieza
                            oConcepto.NoIdentificacion = codigoArticulo;
                            oConcepto.Unidad = "0"; //Por default
                            oConcepto.Descripcion = descripcion;
                            oConcepto.ValorUnitario = Math.Round((valorUni - descuento) / 1.16m, 2);

                            if (descuento == 0)
                            {

                            }
                            else
                            {
                                oConcepto.Descuento = Math.Round(descuento / 1.16m, 2);
                            }

                            oConcepto.Importe = Math.Round((( valorUni - descuento )/ 1.16m) * cantidad, 2);
                            totaltotal += Math.Round(((valorUni - descuento) / 1.16m) * cantidad , 2);

                            ComprobanteConceptoImpuestosTraslado oTraslados = new ComprobanteConceptoImpuestosTraslado();
                            ComprobanteConceptoImpuestos oImpuestos = new ComprobanteConceptoImpuestos();
                            //Aquí pondremos los impuestos
                            oTraslados.Base = Math.Round(importe / 1.16m, 2); //Sin iva
                            oTraslados.Impuesto = c_Impuesto.Item002; //Siempre sera esto
                            oTraslados.TipoFactor = c_TipoFactor.Tasa; // Siempre sera esto
                            oTraslados.TasaOCuota = 0.160000m; //Siempre sera esto
                            oTraslados.Importe = Math.Round((importe / 1.16m) * 0.16m, 2); //Iva

                            ivaTotal = Math.Round(ivaTotal + ((importe / 1.16m) * 0.16m), 2);

                            lstTrasla.Add(oTraslados);
                            oImpuestos.Traslados = lstTrasla.ToArray();
                            lstTrasla.Clear();

                            oConcepto.Impuestos = oImpuestos;

                            lstConceptos.Add(oConcepto);
                            oComprobante.Conceptos = lstConceptos.ToArray();
                        }
                        else
                        {
                            totalNegativo += Convert.ToDecimal(row[0].Cells[17].Value);

                        }
                        conta1++;
                    }
                }
                catch
                {

                }

                ComprobanteImpuestos oImpuestosG = new ComprobanteImpuestos();
                ComprobanteImpuestosTraslado oTrasladosG = new ComprobanteImpuestosTraslado();
                oTrasladosG.Impuesto = c_Impuesto.Item002;
                oTrasladosG.TipoFactor = c_TipoFactor.Tasa;
                oTrasladosG.TasaOCuota = 0.160000m;
                oTrasladosG.Importe = ivaTotal; //TOTAL DE IMPUESTOS

                List<ComprobanteImpuestosTraslado> lstImpTrasla = new List<ComprobanteImpuestosTraslado>();

                lstImpTrasla.Add(oTrasladosG);
                oImpuestosG.Traslados = lstImpTrasla.ToArray();
                oImpuestosG.TotalImpuestosTrasladados = ivaTotal; //TOTAL DE IMPUESTOS
                oComprobante.Impuestos = oImpuestosG;

                oComprobante.SubTotal = Math.Round(totaltotal, 2);
                //MessageBox.Show("Subtotal: " + Math.Round(totaltotal, 2));
                oComprobante.Total = Math.Round(totaltotal + ivaTotal, 2);
                //MessageBox.Show("Subtotal: " + Math.Round(totaltotal + ivaTotal, 2));

                //Crear el XML 
                XML(oComprobante);
                //xsl //Crear la cadena 
                string cadenaOriginal = "";
                string pathxsl = @"\\192.168.0.2\Sistemas\Sistemas\Refacturacion\Sellos\cadenaoriginal_3_3.xslt";
                System.Xml.Xsl.XslCompiledTransform transformador = new System.Xml.Xsl.XslCompiledTransform(true);
                transformador.Load(pathxsl);
                using (StringWriter sw = new StringWriter())
                using (XmlWriter xwo = XmlWriter.Create(sw, transformador.OutputSettings))
                {
                    transformador.Transform(pathXML, xwo);
                    cadenaOriginal = sw.ToString();
                }
                //Sellar el documento
                SelloDigital oselloDigital = new SelloDigital();
                oComprobante.Certificado = oselloDigital.Certificado(pathCer);
                oComprobante.Sello = Test_Mex_Sign_Data(cadenaOriginal);
                //Sobre escribir el xml ya sellado
                XML(oComprobante);

                //Se instancia el WS de Timbrado.
                ServiceReferenceFac.WSCFDI33Client ServicioTimbrado_FEL = new ServiceReferenceFac.WSCFDI33Client();

                //Se instancia la Respuesta del WS de Timbrado.
                ServiceReferenceFac.RespuestaTFD33 RespuestaTimbrado_FEL = new ServiceReferenceFac.RespuestaTFD33();
                ServiceReferenceFac.RespuestaTFD33 RespuestaServicio_FEL = new ServiceReferenceFac.RespuestaTFD33();
                //Se carga el XML desde archivo.
                XmlDocument DocumentoXML = new XmlDocument();
                //La direccion se sustituira dependiendo de donde se leera el XML.
                DocumentoXML.Load(pathxDinamico + diaCarpeta + numOpPathXMLFac + ".xml");

                //Variable string que contiene el contenido del XML.
                string stringXML = null;
                stringXML = DocumentoXML.OuterXml;
                //Timbrar

                // ############## VALIDAR RFC --------------------------------------------------------------
                //ServicioTimbrado_FEL.ValidarRFC("", "", "");

                RespuestaTimbrado_FEL = ServicioTimbrado_FEL.TimbrarCFDI(UsuarioFell, ContraseñaFEll, stringXML, numOp);

                //"DEMO010233001", "Pruebas1a$", "C62D76BA-7E57-7E57-7E57-23288A910663", ""
                //Obteniendo la respuesta se valida que haya sido exitosa.
                if (RespuestaTimbrado_FEL.OperacionExitosa == true)
                {
                    MessageBox.Show("ESTADO DE LA FACTURA: " + RespuestaTimbrado_FEL.Timbre.Estado + System.Environment.NewLine, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    observaciones = RespuestaTimbrado_FEL.Timbre.Estado.ToString();
                    UUIDNUEVO = RespuestaTimbrado_FEL.Timbre.UUID.ToString();

                    DocumentoXML.LoadXml(RespuestaTimbrado_FEL.XMLResultado);
                    DocumentoXML.Save(pathxDinamico + diaCarpeta + numOp + ".xml");

                    //Generar PDF
                    RespuestaServicio_FEL = ServicioTimbrado_FEL.ObtenerPDF(UsuarioFell, ContraseñaFEll, UUIDNUEVO, "");
                    //Guardo el PDF del CFDi.
                    facturaOK = 0;
                    try
                    {
                        File.WriteAllBytes(pathxDinamico + diaCarpeta + numOp + ".pdf", Convert.FromBase64String(RespuestaServicio_FEL.PDFResultado));
                        try
                        {
                            //Enviar pdf por correo(Si es que existe) Crear metodo para enviar
                            EnviarCorreo(pathxDinamico + diaCarpeta + numOp + ".pdf", pathxDinamico + diaCarpeta + numOp + ".xml", to);
                        }
                        catch
                        {
                            MessageBox.Show("NO SE ENVIÓ EL CORREO", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception errorPdf)
                    {
                        MessageBox.Show("ERROR PDF:" + errorPdf, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MessageBox.Show("" + RespuestaServicio_FEL.CodigoRespuesta + System.Environment.NewLine +
                        RespuestaServicio_FEL.MensajeError + System.Environment.NewLine +
                         RespuestaServicio_FEL.MensajeErrorDetallado + System.Environment.NewLine);
                    }


                }
                else
                {
                    MessageBox.Show("ERROR: " + RespuestaTimbrado_FEL.CodigoRespuesta + System.Environment.NewLine + RespuestaTimbrado_FEL.MensajeError + System.Environment.NewLine + RespuestaTimbrado_FEL.MensajeErrorDetallado + System.Environment.NewLine, "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    observaciones = RespuestaTimbrado_FEL.CodigoRespuesta.ToString() + "  " + RespuestaTimbrado_FEL.MensajeErrorDetallado.ToString();
                    facturaOK = 1;
                }
            }
            catch
            {
                MessageBox.Show("NO SE GENERO LA FACTURA", "CASA GUERRERO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                facturaOK = 1;
            }
        }
        
        //CREAR XML
        private void XML(Comprobante oComprobante)
        {
            XmlSerializerNamespaces xmlNamespaces = new XmlSerializerNamespaces();
            xmlNamespaces.Add("cfdi", "http://www.sat.gob.mx/cfd/3");
            xmlNamespaces.Add("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital");
            xmlNamespaces.Add("xsi", "http://www.w3.org/2001/XMLSchema-instance");
            //Serializar el objeto
            XmlSerializer oxmlSerializazar = new XmlSerializer(typeof(Comprobante));
            string sXml = "";
            using (var sww = new XSDToXML.Utils.StringWriterWithEncoding(Encoding.UTF8))
            {
                using (XmlWriter writter = XmlWriter.Create(sww))
                {

                    oxmlSerializazar.Serialize(writter, oComprobante, xmlNamespaces);
                    sXml = sww.ToString();
                }
            }
            //Guardar en un archivo
            System.IO.File.WriteAllText(pathXML, sXml);
        }

        private void checkOtro_CheckedChanged(object sender, EventArgs e)
        {
            if (checkG01.Enabled == false)
            {
                checkG03.Enabled = true;
                checkP01.Enabled = true;
                checkG01.Enabled = true;
                comboOtro.Enabled = false;
                checkP01.Checked = false;
                checkG03.Checked = false;
                checkG01.Checked = false;
            }
            else
            {
                checkG03.Enabled = false;
                checkP01.Enabled = false;
                checkG01.Enabled = false;
                comboOtro.Enabled = true;
                checkP01.Checked = false;
                checkG03.Checked = false;
                checkG01.Checked = false;
            }
        }

        private void checkG03_CheckedChanged(object sender, EventArgs e)
        {
            checkP01.Checked = false;
            checkG01.Checked = false;
        }

        private void checkG01_CheckedChanged(object sender, EventArgs e)
        {
            checkP01.Checked = false;
            checkG03.Checked = false;
        }

        private void checkP01_CheckedChanged(object sender, EventArgs e)
        {
            checkG03.Checked = false;
            checkG01.Checked = false;
        }

        private void fillToolStripButton_Click(object sender, EventArgs e)
        {
            

        }
    }
}
