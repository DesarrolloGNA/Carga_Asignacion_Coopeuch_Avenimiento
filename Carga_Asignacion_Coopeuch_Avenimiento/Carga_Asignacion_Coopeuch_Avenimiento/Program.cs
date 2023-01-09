using System;
using System.Collections.Generic;
using System.IO;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace Carga_Asignacion_Coopeuch_Avenimiento
{
    class Program
    {
        public static List<Asignacion_Coopeuch_Avenimiento> Lista_Asignacion_Coopeuch_Avenimientos = new List<Asignacion_Coopeuch_Avenimiento>();
        static void Main(string[] args)
        {

            /*-------------------------------------------------------------------------*/
            /*                             RUTA DEL ARCHIVO                            */
            /*-------------------------------------------------------------------------*/
            DirectoryInfo di = new DirectoryInfo(@"C:\coopeuch\asignacion\");
            FileInfo[] files = di.GetFiles("*");

            foreach (FileInfo file in files)
            {
                if (file.Name.Contains("Base_Avenimientos"))
                {
                    String Ruta_Archivo = @"C:\coopeuch\asignacion\" + file.Name;
                    Console.WriteLine("Lectura del Archivo: " + file.Name.ToString());
                    /*-------------------------------------------------------------------------*/
                    /*                    llamado al metodo leer archivo                       */
                    /*-------------------------------------------------------------------------*/
                    Leer_Excel(Ruta_Archivo);
                }
            }

            int cantidad = 0;
            foreach (var i in Lista_Asignacion_Coopeuch_Avenimientos)
            {
                Console.WriteLine("Insertando Registro: " + cantidad.ToString());
                /*-------------------------------------------------------------------------*/
                /*                      CARGA A LA BASE DE DATOS                           */
                /*-------------------------------------------------------------------------*/

                string connstring = @"Data Source=192.168.0.77; Initial Catalog=EJFDES; Persist Security Info=True; User ID=sa; Password=Desa2019;";
                using (SqlConnection con = new SqlConnection(connstring))
                {
                    con.Open();
                    string commandString = @"INSERT INTO [dbo].[Coopeuch_Base_Avenimiento]
                        ([Fecha_Carga],[Operación],[Oficina],[Rut socio2],[Nombre socio]
                        ,[Comuna Socio],[Grupo Credito],[Fecha Castigo]
                        ,[Acciones],[Saldo IBS (cierre mes anterior)],[TIPO RECUPERO]
                        ,[Dirección],[Complemento Direccion],[Nombre Comuna],[Celular]
                        ,[Telefono Primario],[Telefono Secundario],[Direccion Email]
                        ,[tipo deuda],[INTERES MORA],[TASA],[SCRIPT])
                        VALUES(@FECHA_CARGA,@OPERACION
                        ,@OFICINA,@RUT_SOCIO2,@NOMBRE_SOCIO,@COMUNA_SOCIO
                        ,@GRUPO_CREDITO,@FECHA_CASTIGO,@ACCIONES,@SALDO_IBS_CIERRE_MES_ANTERIOR
                        ,@TIPO_RECUPERO,@DIRECCION,@COMPLEMENTO_DIRECCION,@NOMBRE_COMUNA
                        ,@CELULAR,@TELEFONO_PRIMARIO,@TELEFONO_SECUNDARIO,@DIRECCION_EMAIL
                        ,@TIPO_DEUDA,@INTERES_MORA,@TASA,NULL)";

                    SqlCommand cmd = new SqlCommand(commandString, con);
                    cmd.Parameters.AddWithValue("@FECHA_CARGA", i.FECHA_CARGA);
                    cmd.Parameters.AddWithValue("@OPERACION", i.OPERACION);
                    cmd.Parameters.AddWithValue("@OFICINA", i.OFICINA);
                    cmd.Parameters.AddWithValue("@RUT_SOCIO2", i.RUT_SOCIO2);
                    cmd.Parameters.AddWithValue("@NOMBRE_SOCIO", i.NOMBRE_SOCIO);
                    cmd.Parameters.AddWithValue("@COMUNA_SOCIO", i.COMUNA_SOCIO);
                    //cmd.Parameters.AddWithValue("@ANIO_CASTIGO", i.ANIO_CASTIGO);
                    cmd.Parameters.AddWithValue("@GRUPO_CREDITO", i.GRUPO_CREDITO);
                    cmd.Parameters.AddWithValue("@FECHA_CASTIGO", i.FECHA_CASTIGO);
                    cmd.Parameters.AddWithValue("@ACCIONES", i.ACCIONES);
                    cmd.Parameters.AddWithValue("@SALDO_IBS_CIERRE_MES_ANTERIOR", i.SALDO_IBS_CIERRE_MES_ANTERIOR);
                    cmd.Parameters.AddWithValue("@TIPO_RECUPERO", i.TIPO_RECUPERO);
                    cmd.Parameters.AddWithValue("@DIRECCION", i.DIRECCION);
                    cmd.Parameters.AddWithValue("@COMPLEMENTO_DIRECCION", i.COMPLEMENTO_DIRECCION);
                    cmd.Parameters.AddWithValue("@NOMBRE_COMUNA", i.NOMBRE_COMUNA);
                    cmd.Parameters.AddWithValue("@CELULAR", i.CELULAR);
                    cmd.Parameters.AddWithValue("@TELEFONO_PRIMARIO", i.TELEFONO_PRIMARIO);
                    cmd.Parameters.AddWithValue("@TELEFONO_SECUNDARIO", i.TELEFONO_SECUNDARIO);
                    cmd.Parameters.AddWithValue("@DIRECCION_EMAIL", i.DIRECCION_EMAIL);
                    cmd.Parameters.AddWithValue("@TIPO_DEUDA", i.TIPO_DEUDA);
                    cmd.Parameters.AddWithValue("@INTERES_MORA", i.INTERES_MORA);
                    cmd.Parameters.AddWithValue("@TASA", i.TASA);

                    cmd.ExecuteNonQuery();
                    con.Close();
                    cantidad++;
                }
            }
            Console.ReadKey();
            Lista_Asignacion_Coopeuch_Avenimientos = null;
            GC.Collect();
        }

        private static void Leer_Excel(string ruta)
        {
            int contador = 0;
            string Campo = "";
            DateTime Fecha_Actual = DateTime.Now;
            /*--------------------------------------------------------------------------------*/
            /*                              LECTURA DE ARCHIVO                                */
            /*--------------------------------------------------------------------------------*/
            try
            {
                using (var stream = File.Open(ruta, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                    {
                        FallbackEncoding = Encoding.UTF8,
                        LeaveOpen = false,
                        AnalyzeInitialCsvRows = 0,
                    }))
                    {
                        do
                        {
                            //recorro el excel
                            while (reader.Read())
                            {
                                //omito cabecera
                                if (contador > 1)
                                {
                                    Asignacion_Coopeuch_Avenimiento Input = new Asignacion_Coopeuch_Avenimiento();

                                    Campo = " FECHA_CARGA";
                                    Input.FECHA_CARGA = Fecha_Actual;
                                    Campo = " OPERACION";
                                    Input.OPERACION = reader.GetValue(0) != null ? reader.GetValue(0).ToString() : "";
                                    Campo = " OFICINA";
                                    Input.OFICINA = reader.GetValue(10) != null ? reader.GetValue(10).ToString() : "";
                                    Campo = " RUT_SOCIO2";
                                    Input.RUT_SOCIO2 = reader.GetValue(12) != null ? int.Parse(reader.GetValue(12).ToString()) : 0;
                                    Campo = " NOMBRE_SOCIO";
                                    Input.NOMBRE_SOCIO = reader.GetValue(14) != null ? reader.GetValue(14).ToString() : "";
                                    Campo = " COMUNA_SOCIO";
                                    Input.COMUNA_SOCIO = reader.GetValue(15) != null ? reader.GetValue(15).ToString() : "";
                                    Campo = " ANIO_CASTIGO";
                                    //Input.ANIO_CASTIGO = reader.GetValue(18) != null ? int.Parse(reader.GetValue(18).ToString()) : 0;
                                    Campo = " GRUPO_CREDITO";
                                    Input.GRUPO_CREDITO = reader.GetValue(19) != null ? reader.GetValue(19).ToString() : "";
                                    Campo = " FECHA_CASTIGO";
                                    Input.FECHA_CASTIGO = reader.GetValue(16) != null ? DateTime.Parse(reader.GetValue(16).ToString()) : DateTime.Parse("1900-01-01");
                                    Campo = " ACCIONES";
                                    Input.ACCIONES = reader.GetValue(24) != null ? reader.GetValue(24).ToString() : "";
                                    Campo = " SALDO_IBS_CIERRE_MES_ANTERIOR";
                                    Input.SALDO_IBS_CIERRE_MES_ANTERIOR = reader.GetValue(29) != null ? int.Parse(reader.GetValue(29).ToString()) : 0;
                                    Campo = " TIPO_RECUPERO";
                                    Input.TIPO_RECUPERO = reader.GetValue(37) != null ? reader.GetValue(37).ToString() : "";
                                    Campo = " DIRECCION";
                                    Input.DIRECCION = reader.GetValue(38) != null ? reader.GetValue(38).ToString() : "";
                                    Campo = " COMPLEMENTO_DIRECCION";
                                    Input.COMPLEMENTO_DIRECCION = reader.GetValue(39) != null ? reader.GetValue(39).ToString() : "";
                                    Campo = " NOMBRE_COMUNA";
                                    Input.NOMBRE_COMUNA = reader.GetValue(40) != null ? reader.GetValue(40).ToString() : "";
                                    Campo = " CELULAR";
                                    Input.CELULAR = reader.GetValue(41) != null ? reader.GetValue(41).ToString() : "";
                                    Campo = " TELEFONO_PRIMARIO";
                                    Input.TELEFONO_PRIMARIO = reader.GetValue(42) != null ? reader.GetValue(42).ToString() : "";
                                    Campo = " TELEFONO_SECUNDARIO";
                                    Input.TELEFONO_SECUNDARIO = reader.GetValue(43) != null ? reader.GetValue(43).ToString() : "";
                                    Campo = " DIRECCION_EMAIL";
                                    Input.DIRECCION_EMAIL = reader.GetValue(44) != null ? reader.GetValue(44).ToString() : "";
                                    Campo = " TIPO_DEUDA";
                                    Input.TIPO_DEUDA = reader.GetValue(47) != null ? reader.GetValue(47).ToString() : "";
                                    Campo = " INTERES_MORA";
                                    Input.INTERES_MORA = reader.GetValue(65) != null ? int.Parse(reader.GetValue(65).ToString()) : 0;
                                    Campo = " TASA";
                                    Input.TASA = reader.GetValue(66) != null ? float.Parse(reader.GetValue(66).ToString()) : 0;

                                    Lista_Asignacion_Coopeuch_Avenimientos.Add(Input);
                                }
                                contador++;
                            }
                        } while (reader.NextResult());
                    }
                }
            }
            catch (Exception e)
            {
                string mensaje = "Campo: " + Campo + " contador: " + contador + " error: " + e;
                Console.WriteLine(mensaje);
            }
        }

        public class Asignacion_Coopeuch_Avenimiento
        {
            public DateTime FECHA_CARGA { get; set; }
            public String OPERACION { get; set; }
            public String OFICINA { get; set; }
            public int RUT_SOCIO2 { get; set; }
            public String NOMBRE_SOCIO { get; set; }
            public String COMUNA_SOCIO { get; set; }
            public int ANIO_CASTIGO { get; set; }
            public String GRUPO_CREDITO { get; set; }
            public DateTime FECHA_CASTIGO { get; set; }
            public String ACCIONES { get; set; }
            public int SALDO_IBS_CIERRE_MES_ANTERIOR { get; set; }
            public String TIPO_RECUPERO { get; set; }
            public String DIRECCION { get; set; }
            public String COMPLEMENTO_DIRECCION { get; set; }
            public String NOMBRE_COMUNA { get; set; }
            public String CELULAR { get; set; }
            public String TELEFONO_PRIMARIO { get; set; }
            public String TELEFONO_SECUNDARIO { get; set; }
            public String DIRECCION_EMAIL { get; set; }
            public String TIPO_DEUDA { get; set; }
            public int INTERES_MORA { get; set; }
            public float TASA { get; set; }

        }
    }
}
