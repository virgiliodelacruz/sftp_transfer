/*
 * 
 * 
DESARROLLO:			sftp_transfer
OBJETIVO:			Este desarrollo realiza lo siguiente en ese orden
 *                  1) Borra los archivos xml que esten en la carpeta donde este el ejectuable sftp_transfer.exe
 *                  2) Copia archivos xml de la carpeta C:/subirftp a la carpeta actual (donde esta el ejecutable), 
 *                     solo se copian aquellos que fueron generados o modificados el dia de hoy
 *                  3) Sube archivos xml que se encuentran en la carpeta actual al servidor ftp
 *                  4) Borra los archivos xml que esten en la carpeta donde este el ejectuable sftp_transfer.exe
 * 
AUTORES:			Virgilio De la Cruz
FECHA:		20150521
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Globalization;
using System.Configuration;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using Renci.SshNet;
using System.Xml;
using System.Text;

namespace sftp_transfer
{
    class Program
    {
        //Trae las variables del archivo de configuración
        public static String rutallaveprivada = ConfigurationManager.AppSettings["rutallaveprivada"].ToString();
        public static String rutaasubirftp = ConfigurationManager.AppSettings["rutaasubirftp"].ToString();
        public static String hostnamesftp = ConfigurationManager.AppSettings["hostnamesftp"].ToString();
        public static String portsftp = ConfigurationManager.AppSettings["portsftp"].ToString();
        public static String usernamesftp = ConfigurationManager.AppSettings["usernamesftp"].ToString();
        public static String rutacarpetaasubir = ConfigurationManager.AppSettings["rutacarpetaasubir"].ToString();
        public static String hostbasededatos = ConfigurationManager.AppSettings["hostbasededatos"].ToString(); //Propiedad de la firma del documento PDF
        public static String portbasededatos = ConfigurationManager.AppSettings["portbasededatos"].ToString(); //Configuración de layout de firma
        public static String servicenamebasededatos = ConfigurationManager.AppSettings["servicenamebasededatos"].ToString(); //Configuración de layout de firma
        public static String useridbasededatos = ConfigurationManager.AppSettings["useridbasededatos"].ToString(); //Configuración de layout de firma
        public static String passwordbasededatos = ConfigurationManager.AppSettings["passwordbasededatos"].ToString(); //Configuración de layout de firma
        public static System.IO.StreamWriter filebit;
        public static void borraarchivosxml()
        {
            //Traigo la lista de archivos XML de la carpeta actual (donde esta el ejecutable)
            string[] listaxml = Directory.GetFiles(".", "*.xml");
            //Recorro la lista de nombres de archivos
            foreach (string f in listaxml)
            {
                //Los borro
                File.Delete(f);
            }

            
        }
        
        public static void copiaarchivos()
        {
            //Busco archivos xml que esten en la carpeta C:/subirftp
            foreach (string path in Directory.GetFiles(rutacarpetaasubir, "*.xml", SearchOption.AllDirectories))
            {
                //Obtengo fecha de creación y de última modificación de cada archivo
                DateTime d = File.GetCreationTime(path);
                DateTime d1 = File.GetLastWriteTime(path);
                //Si la fecha de creación o modificación es el dia de hoy ...
                if ((d.ToShortDateString() == DateTime.Today.ToShortDateString())  ||  (d1.ToShortDateString() == DateTime.Today.ToShortDateString()))
                {
                    //Copio el archivo de la carpeta C:/subirftp a la carpeta actual
                    File.Copy(path, path.Replace(rutacarpetaasubir, ""));
                }
            }
        }

        public static void SubeArchivoXMLaFTP(string archivo)
        {
            //variables de configuración para conectarse al ftp
            var keyFile = new PrivateKeyFile(rutallaveprivada);
            var keyFiles = new[] { keyFile };
            var methods = new List<AuthenticationMethod>();
            methods.Add(new PrivateKeyAuthenticationMethod(usernamesftp, keyFiles));

            var con = new ConnectionInfo(hostnamesftp, Convert.ToInt32(portsftp), usernamesftp, methods.ToArray());
            using (var client = new SftpClient(con))
            {
                //me conecto al sftp...
                try
                {
                    client.Connect();
                    client.ChangeDirectory(rutaasubirftp);
                }
                catch (Exception e)
                {
                    //si hubiera un error
                    filebit.WriteLine("Error al conectarse al servidor " + hostnamesftp + "Info detallada del error=" + e.Data + "-" + e.Message);
                    Console.WriteLine("Error al conectarse al servidor " + hostnamesftp + "Info detallada del error=" + e.Data  +"-"+e.Message);
                }
                //me cambio a la carpeta deseada
              

                //busco loas archivos xml de la carpeta actual
                   using (
                        //leo el archivo
                        var uplfileStream = System.IO.File.OpenRead(archivo))
                    {
                        try
                        {
                            //lo subo
                            client.UploadFile(uplfileStream, archivo, true);
                           // string name = archivo.Substring(2, archivo.Length - 2);
                            //mando un mensaje a la pantalla
                            filebit.WriteLine("Archivo " + archivo + " subido correctamente");
                            System.Console.WriteLine("Archivo " + archivo + " subido correctamente");
                        }
                        catch (Exception e)
                        {
                            //si hubiera un error
                            filebit.WriteLine("Error al subir el archivo " + archivo + " error=" + e.Data + " mesj=" + e.Message);
                            System.Console.WriteLine("Error al subir el archivo " + archivo + " error=" + e.Data + " mesj=" + e.Message);
                        }
                    }
                              //se desconecta
                client.Disconnect();
            }

        }
        public static void SubeArchivos()
        {
            //variables de configuración para conectarse al ftp
            var keyFile = new PrivateKeyFile(rutallaveprivada);
            var keyFiles = new[] { keyFile };
            var methods = new List<AuthenticationMethod>();
            methods.Add(new PrivateKeyAuthenticationMethod(usernamesftp, keyFiles));

            var con = new ConnectionInfo(hostnamesftp, Convert.ToInt32(portsftp), usernamesftp, methods.ToArray());
            using (var client = new SftpClient(con))
            {
               //me conecto al sftp...
                client.Connect();
                //me cambio a la carpeta deseada
                client.ChangeDirectory(rutaasubirftp);

                //busco loas archivos xml de la carpeta actual
                foreach (string path in Directory.GetFiles(".", "*.xml", SearchOption.TopDirectoryOnly))
                {
                    using (
                    //leo el archivo
                        var uplfileStream = System.IO.File.OpenRead(path))
                    {
                        try
                        {
                            //lo subo
                            client.UploadFile(uplfileStream, path, true);
                            string name = path.Substring(2, path.Length - 2 );
                            //mando un mensaje a la pantalla
                            filebit.WriteLine("Archivo " + name + " subido correctamente");
                            Console.WriteLine("Archivo " + name + " subido correctamente");
                        }
                        catch (Exception e)
                        {
                            //si hubiera un error
                            filebit.WriteLine("Error al subir el archivo " + path + " error=" + e.Data + " mesj=" + e.Message);
                            System.Console.WriteLine("Error al subir el archivo " + path +" error="+ e.Data + " mesj=" + e.Message);
                        }
                    }
                }
               //se desconecta
                client.Disconnect();
            }
            
        }
        public static string getStringConexion()
        {
            string conexion = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)" +
                                            "(HOST=" + hostbasededatos + ")(PORT=" + portbasededatos + "))" +
                                            "(CONNECT_DATA=(SERVER=dedicated)(SERVICE_NAME=" + servicenamebasededatos + ")));" +
                                            "user id=" + useridbasededatos + ";password=" + passwordbasededatos + ";";
            return conexion;
        }
        public static void ActualizaBDUploaded(string archivo)
        {
            OracleConnection oracleConn = new OracleConnection();
            oracleConn.ConnectionString = getStringConexion();
            //string ruta = @"C:\subirftp\" + archivo ;
           // String SourceLoc = @"C:\subirftp\" + archivo;


            //FileStream fs = new FileStream(SourceLoc, FileMode.Open, FileAccess.Read);
            //byte[] XMLData = new byte[fs.Length];
            //fs.Read(XMLData, 0, System.Convert.ToInt32(fs.Length));

            //Close the File Stream
            //fs.Close();


            try
            {
                oracleConn.Open();
                String block = " BEGIN " +
                " UPDATE anahuac.GWBAXML SET GWBAXML_UPLOADED ='Y' WHERE GWBAXML_FILE_NAME ='" + archivo + "'; " +
                " commit; "+
                " END; ";
            
                OracleCommand cmd = new OracleCommand();
                cmd.CommandText = block;
                cmd.Connection = oracleConn;
                cmd.CommandType = CommandType.Text;               
                //OracleParameter param = cmd.Parameters.Add("blobtodb", OracleDbType.Blob);
                //param.Direction = ParameterDirection.Input;

               // param.Value = XMLData;
                cmd.ExecuteNonQuery();
                filebit.WriteLine(archivo + " marcado como subido en la bd...");
                System.Console.WriteLine(archivo + " marcado como subido en la bd...");

            }

            catch (Exception ex)
            {
                filebit.WriteLine("Exception hay trans: {0}", ex.ToString());
                System.Console.WriteLine("Exception hay trans: {0}", ex.ToString());

            }
          //  return "ok";
        }
        public static void insertaxmlenbd(string archivo,string tipo)
        {
            OracleConnection oracleConn = new OracleConnection();
            oracleConn.ConnectionString = getStringConexion();
            //string ruta = @"C:\subirftp\" + archivo ;

            
           // String SourceLoc = @"C:\subirftp\" + archivo;
            String SourceLoc = rutacarpetaasubir + archivo;

            FileStream fs = new FileStream(SourceLoc, FileMode.Open, FileAccess.Read);
            byte[] XMLData = new byte[fs.Length];
            fs.Read(XMLData, 0, System.Convert.ToInt32(fs.Length));

            //Close the File Stream
            fs.Close();


            try
            {
                oracleConn.Open();
                String block = " BEGIN " +
                " INSERT INTO anahuac.GWBAXML VALUES (seq_gwbaxml.nextval,'" + archivo + "' ,:1,'"+tipo+"','N',SYSDATE,USER); " +
                " commit; " +
                " END; ";

                OracleCommand cmd = new OracleCommand();
                cmd.CommandText = block;
                cmd.Connection = oracleConn;
                cmd.CommandType = CommandType.Text;
                OracleParameter param = cmd.Parameters.Add("blobtodb", OracleDbType.Blob);
                param.Direction = ParameterDirection.Input;

                param.Value = XMLData;
                cmd.ExecuteNonQuery();
                filebit.WriteLine(archivo + " insertado en la base");
                System.Console.WriteLine(archivo + " insertado en la base");

            }

            catch (Exception ex)
            {
                filebit.WriteLine("Exception hay trans: {0}", ex.ToString());
                System.Console.WriteLine("Exception hay trans: {0}", ex.ToString());

            }
            //  return "ok";
        }
        public static string generaxmldeconsulta(string tipo)
        {

         
            // FileInfo file = new FileInfo(@"C:\subirftp\"+tipo+".sql");
            FileInfo file = new FileInfo(rutacarpetaasubir + tipo + ".sql");
            string queryString = file.OpenText().ReadToEnd();
            OracleConnection oracleConn = new OracleConnection();
            oracleConn.ConnectionString = getStringConexion();
            string dia = Convert.ToString(DateTime.Today.Day);
            string mes = Convert.ToString(DateTime.Today.Month);
            string anio = Convert.ToString(DateTime.Today.Year);
            string hora = Convert.ToString(DateTime.Now.Hour);
            string minuto = Convert.ToString(DateTime.Now.Minute);
            string segundo = Convert.ToString(DateTime.Now.Second);
            if (dia.Length == 1)
                dia = "0" + dia;
            if (mes.Length == 1)
                mes = "0" + mes;
            if (anio.Length == 1)
                anio = "0" + anio;
            if (hora.Length == 1)
                hora = "0" + hora;
            if (minuto.Length == 1)
                minuto = "0" + minuto;
            if (segundo.Length == 1)
                segundo = "0" + segundo;

            string filename = "UAMS-" + tipo + "-" + dia + mes + anio + "-" + hora + minuto + segundo;
           // XmlWriter xmlWriter = XmlWriter.Create(@"C:\subirftp\" + filename + ".xml");
            XmlWriter xmlWriter = XmlWriter.Create(rutacarpetaasubir + filename + ".xml"); 
            xmlWriter.WriteStartDocument();

            try
            {

                oracleConn.Open();
                OracleCommand Cmd2 = new OracleCommand(queryString, oracleConn);
                OracleDataReader reader = Cmd2.ExecuteReader();
                xmlWriter.WriteStartElement(tipo);
                xmlWriter.WriteAttributeString("university", "UAMS");
                try
                {

                    while (reader.Read())
                    {

                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                           
                            xmlWriter.WriteStartElement(reader.GetName(i));
                            xmlWriter.WriteString(Convert.ToString(reader[reader.GetName(i)]));
                            xmlWriter.WriteEndElement();
                        }
                    }
                    xmlWriter.WriteEndElement();
                    xmlWriter.WriteEndDocument();
                    xmlWriter.Close();
                    
                }
                finally
                {
                    
                    reader.Close();
                }

            }
            catch (Exception ex)
            {
                 filebit.WriteLine("Error: {0}", ex.ToString());
                System.Console.WriteLine("Error: {0}", ex.ToString());
                
            }
            xmlWriter.Close();
            xmlWriter.Close();
            oracleConn.Close();
            return filename + ".xml";
        }
        static void Proceso(string tipo)
        {
            
            borraarchivosxml();
            string archivo;
            filebit.WriteLine("Generando xml " + tipo + " ...");
            System.Console.WriteLine("Generando xml " + tipo +" ...");
            archivo = generaxmldeconsulta(tipo);
            filebit.WriteLine("Insertando archivo xml " + archivo + " en BD..");
            System.Console.WriteLine("Insertando archivo xml " + archivo + " en BD..");
            insertaxmlenbd(archivo,tipo);
            copiaarchivos();
            //SubeArchivos();
            filebit.WriteLine("Subiendo archivo " + archivo + " a ftp...");
            System.Console.WriteLine("Subiendo archivo " + archivo + " a ftp...");
            SubeArchivoXMLaFTP(archivo);
            ActualizaBDUploaded(archivo);
            borraarchivosxml();
           
        }
        static void Main(string[] args)
        {
            //la ejecución de los pasos descritos en los comentarios al inicio
            //generaxmldeconsulta1();
           // System.Console.WriteLine("Borrando xml..");
            string dia = Convert.ToString(DateTime.Today.Day);
            string mes = Convert.ToString(DateTime.Today.Month);
            string anio = Convert.ToString(DateTime.Today.Year);
            string hora = Convert.ToString(DateTime.Now.Hour);
            string minuto = Convert.ToString(DateTime.Now.Minute);
            string segundo = Convert.ToString(DateTime.Now.Second);
            if (dia.Length == 1)
                dia = "0" + dia;
            if (mes.Length == 1)
                mes = "0" + mes;
            if (anio.Length == 1)
                anio = "0" + anio;
            if (hora.Length == 1)
                hora = "0" + hora;
            if (minuto.Length == 1)
                minuto = "0" + minuto;
            if (segundo.Length == 1)
                segundo = "0" + segundo;           
           filebit = new System.IO.StreamWriter(rutacarpetaasubir + "bitacora-" + dia + mes + anio + "-" + hora + minuto + segundo + ".txt");
           filebit.WriteLine("----INICIO bitacora-" + dia + mes + anio + "-" + hora + minuto + segundo+"-------");
            Proceso("Student");
            Proceso("Applicant");
            Proceso("Enrollment");
            Proceso("Section");
            filebit.WriteLine("----FIN bitacora-" + dia + mes + anio + "-" + hora + minuto + segundo + "-------");
            filebit.Close();
        }
    }
}
