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
                        ,[Comuna Socio],[Año Castigo],[Grupo Credito],[Fecha Castigo]
                        ,[Acciones],[Saldo IBS (cierre mes anterior)],[TIPO RECUPERO]
                        ,[Dirección],[Complemento Direccion],[Nombre Comuna],[Celular]
                        ,[Telefono Primario],[Telefono Secundario],[Direccion Email]
                        ,[tipo deuda],[INTERES MORA],[TASA],[SCRIPT])
                        VALUES(@ID_COOPEUCH_BASE_AVENIMIENTO,@FECHA_CARGA,@OPERACION
                        ,@OFICINA,@RUT_SOCIO2,@NOMBRE_SOCIO,@COMUNA_SOCIO,@ANIO_CASTIGO
                        ,@GRUPO_CREDITO,@FECHA_CASTIGO,@ACCIONES,@SALDO_IBS_CIERRE_MES_ANTERIOR
                        ,@TIPO_RECUPERO,@DIRECCION,@COMPLEMENTO_DIRECCION,@NOMBRE_COMUNA
                        ,@CELULAR,@TELEFONO_PRIMARIO,@TELEFONO_SECUNDARIO,@DIRECCION_EMAIL
                        ,@TIPO_DEUDA,@INTERES_MORA,@TASA,@SCRIPT)";

                    SqlCommand cmd = new SqlCommand(commandString, con);
                    cmd.Parameters.AddWithValue("@FECHA_CARGA", i.FECHA_CARGA);
                    cmd.Parameters.AddWithValue("@OPERACION", i.OPERACION);
                    cmd.Parameters.AddWithValue("@OFICINA", i.OFICINA);
                    cmd.Parameters.AddWithValue("@RUT_SOCIO2", i.RUT_SOCIO2);
                    cmd.Parameters.AddWithValue("@NOMBRE_SOCIO", i.NOMBRE_SOCIO);
                    cmd.Parameters.AddWithValue("@COMUNA_SOCIO", i.COMUNA_SOCIO);
                    cmd.Parameters.AddWithValue("@ANIO_CASTIGO", i.ANIO_CASTIGO);
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

        private static void Leer_Excel(string ruta_Archivo)
        {
            int contador = 0;
            string Campo = "";
            DateTime Fecha_Actual = DateTime.Now;
            /*--------------------------------------------------------------------------------*/
            /*                              LECTURA DE ARCHIVO                                */
            /*--------------------------------------------------------------------------------*/
            try
            {

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
            public int TASA { get; set; }

        }
    }
}
