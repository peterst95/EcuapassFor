using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.Windows.Forms;
using Spire.Xls;
using System.Runtime.InteropServices;

namespace EcuapassForm
{
    public partial class Form1 : Form
    {



        string cs = ConfigurationManager.ConnectionStrings["omarsaConnectionStringProd"].ConnectionString;
        string cs2 = ConfigurationManager.ConnectionStrings["omarsaConnectionStringIntranet"].ConnectionString;

        public Form1()
        {
            InitializeComponent();
        }


        private static string[] _knownFolderGuids = new string[]
 {
        "{56784854-C6CB-462B-8169-88E350ACB882}", // Contacts
        "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}", // Desktop
        "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}", // Documents
        "{374DE290-123F-4565-9164-39C4925E467B}", // Downloads
        "{1777F761-68AD-4D8A-87BD-30B759FA33DD}", // Favorites
        "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}", // Links
        "{4BD8D571-6D19-48D3-BE97-422220080E43}", // Music
        "{33E28130-4E1E-4676-835A-98395C3BC3BB}", // Pictures
        "{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}", // SavedGames
        "{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}", // SavedSearches
        "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}", // Videos
 };

        /// <summary>
        /// Gets the current path to the specified known folder as currently configured. This does
        /// not require the folder to be existent.
        /// </summary>
        /// <param name="knownFolder">The known folder which current path will be returned.</param>
        /// <returns>The default path of the known folder.</returns>
        /// <exception cref="System.Runtime.InteropServices.ExternalException">Thrown if the path
        ///     could not be retrieved.</exception>
        public static string GetPath(KnownFolder knownFolder)
        {
            return GetPath(knownFolder, false);
        }

        /// <summary>
        /// Gets the current path to the specified known folder as currently configured. This does
        /// not require the folder to be existent.
        /// </summary>
        /// <param name="knownFolder">The known folder which current path will be returned.</param>
        /// <param name="defaultUser">Specifies if the paths of the default user (user profile
        ///     template) will be used. This requires administrative rights.</param>
        /// <returns>The default path of the known folder.</returns>
        /// <exception cref="System.Runtime.InteropServices.ExternalException">Thrown if the path
        ///     could not be retrieved.</exception>
        public static string GetPath(KnownFolder knownFolder, bool defaultUser)
        {
            return GetPath(knownFolder, KnownFolderFlags.DontVerify, defaultUser);
        }

        private static string GetPath(KnownFolder knownFolder, KnownFolderFlags flags,
            bool defaultUser)
        {
            int result = SHGetKnownFolderPath(new Guid(_knownFolderGuids[(int)knownFolder]),
                (uint)flags, new IntPtr(defaultUser ? -1 : 0), out IntPtr outPath);
            if (result >= 0)
            {
                string path = Marshal.PtrToStringUni(outPath);
                Marshal.FreeCoTaskMem(outPath);
                return path;
            }
            else
            {
                throw new ExternalException("Unable to retrieve the known folder path. It may not "
                    + "be available on this system.", result);
            }
        }

        [DllImport("Shell32.dll")]
        private static extern int SHGetKnownFolderPath(
            [MarshalAs(UnmanagedType.LPStruct)]Guid rfid, uint dwFlags, IntPtr hToken,
            out IntPtr ppszPath);

        [Flags]
        private enum KnownFolderFlags : uint
        {
            SimpleIDList = 0x00000100,
            NotParentRelative = 0x00000200,
            DefaultPath = 0x00000400,
            Init = 0x00000800,
            NoAlias = 0x00001000,
            DontUnexpand = 0x00002000,
            DontVerify = 0x00004000,
            Create = 0x00008000,
            NoAppcontainerRedirection = 0x00010000,
            AliasOnly = 0x80000000
        }
 

    /// <summary>
    /// Standard folders registered with the system. These folders are installed with Windows Vista
    /// and later operating systems, and a computer will have only folders appropriate to it
    /// installed.
    /// </summary>
    public enum KnownFolder
    {
        Contacts,
        Desktop,
        Documents,
        Downloads,
        Favorites,
        Links,
        Music,
        Pictures,
        SavedGames,
        SavedSearches,
        Videos
    }


    private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();

            SqlDataAdapter adapt2;
            SqlDataAdapter adapt3;
            SqlDataAdapter adapt4;
            SqlDataAdapter adapt5;
            SqlDataAdapter adapt6;
            SqlDataAdapter adapt7;
            SqlDataAdapter adapt8;
            DataTable dt2;
            DataTable dt3;
            DataTable dt4;
            DataTable dt5;
            DataTable dt6;
            DataTable dt7;
            DataTable dt8;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection con;
            SqlDataAdapter adapt1;
            DataTable dt1;

            dt1 = new DataTable();
            con = new SqlConnection(cs);
            con.Open();
            adapt1 = new SqlDataAdapter("SELECT F.NUMERO_FACTURA,F.ESTADO_CALIDAD AS CALIDAD,F.ESTADO_SANITARIO AS  SANITARIO FROM[omarsa].[dbo].TBL_FACTURAS_CONTROL F WHERE F.ESTADO_CALIDAD = 0 OR F.ESTADO_SANITARIO = 0  ", con);
            adapt1.Fill(dt1);

            if (dt1.Rows.Count > 0)
            {
                GridFacturas.DataSource = dt1;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conintr;
            SqlConnection conprod;
            conprod = new SqlConnection(cs);
           // conintr = new SqlConnection(cs2);

           
            SqlDataAdapter adapt2;
            SqlDataAdapter adapt3;
            SqlDataAdapter adapt4;
            SqlDataAdapter adapt5;
            SqlDataAdapter adapt6;
            SqlDataAdapter adapt7;
            DataTable dt2;
            DataTable dt3;
            DataTable dt4;

           

            String  id_emb;
            String  numero_factura;

            dt2 = new DataTable();
            conprod.Open();

            adapt2 = new SqlDataAdapter("SELECT TOP 1 F.ID,f.NUMERO_FACTURA FROM[omarsa].[dbo].TBL_FACTURAS_CONTROL F WHERE F.ESTADO_CALIDAD = 0 OR F.ESTADO_SANITARIO = 0  ", conprod);
            adapt2.SelectCommand.CommandTimeout = 0;
            adapt2.Fill(dt2);

            //obtengo id del plan Embarque
            id_emb = dt2.Rows[0]["id"].ToString();
            numero_factura = dt2.Rows[0]["numero_factura"].ToString();

            dt3 = new DataTable();



            

            ////// CONSULTA DE CERTIFICADOS DE CALIDAD
            //conprod.Open();          

            adapt3 = new SqlDataAdapter("use omarsa " +
                                        "IF OBJECT_ID('tempdb..#TMP_ECUAPASS') IS NOT NULL   BEGIN   DROP TABLE #TMP_ECUAPASS END " +
                                        "IF OBJECT_ID('tempdb..#calidad') IS NOT NULL        BEGIN   DROP TABLE #calidad END " +
                                        "CREATE TABLE #TMP_ECUAPASS ([hc] [varchar](max) ,[prdt_nm] [varchar](max) ,idprod uniqueidentifier,[prdt_stn] [varchar](max) ,[pck_qt] [decimal](18, 6) ,[pck_ut] [varchar](max) ,[prdt_nwt] [decimal](18, 6) ,[nwt_ut] [varchar](max) ,[pdtn_de] [varchar](max) ,[metrica] [varchar](max) ,[lote] [varchar](max) ,[lot_no] [varchar](max) ,[tipo] [varchar](max)) " +
                                        "INSERT INTO #TMP_ECUAPASS " +
                                        "EXEC [dbo].[SP_CONTROL_CODIGOS_EMBARQUE_REAL_ECUAPASS_CHINA] '" + id_emb + "'  " +
                                        "CREATE TABLE #calidad([Subpartida Arancelaria] varchar(max),[Nombre de Producto] varchar(max), [Nombre de Especie de producto] varchar (max),[Presentacion de Producto] varchar(max),[Tipo de Análisis] varchar(max),[Cantidad de Producto]  varchar(max),[Unidad de Cantidad de Producto]  varchar(max),[Peso Neto de Producto] varchar(max),[Unidad de Peso Neto de Producto] varchar (max),[Código de Lote] varchar (max)  ) "+
                                        "INSERT INTO #calidad	VALUES ('hc','prdt_nm','prdt_spc_nm','prdt_smt_frm_inf','anls_type_nm','prdt_qt','prdt_qt_ut','prdt_nwt','prdt_nwt_ut','lot_cd') "+
                                        "INSERT INTO #calidad	VALUES ('-','-','-','-','-','-','COM_0029','-','COM_0029','-') "+
                                        "INSERT INTO #calidad "+
                                        "SELECT "+
                                        "E.hc, "+
                                        "'CAMARON CONGELADO', "+
                                        "E.prdt_stn, " +
                                        "U.DVINGLES AS descripcion, " +
                                        "'MANCHA BLANCA (WSSV)' TIPO, " +
                                        "CAST(SUM(E.pck_qt) AS INT) pck_qt, " +
                                        "pck_ut, " +
                                        "CAST(SUM(prdt_nwt) AS INT) prdt_nwt, " +
                                        "nwt_ut, " +
                                        "e.lot_no " +
                                        "FROM #TMP_ECUAPASS E " +
                                        "LEFT JOIN PRODUCTO P    ON E.IDPROD = P.ID " +
                                        "LEFT JOIN UD_PRODUCTO U ON U.ID = P.BOEXTENSION_ID " +
                                        "group by e.hc, prdt_nm, u.DVINGLES, prdt_stn, pck_ut, nwt_ut, e.lot_no " +
                                        "SELECT * FROM #calidad "  , conprod);
                                        adapt3.SelectCommand.CommandTimeout = 0;
                                        adapt3.Fill(dt3);
                                        dataGridView1.DataSource = dt3;

            //Descarga archivo Calidad 
            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            book.Worksheets[0].InsertDataTable(this.dataGridView1.DataSource as DataTable, true, 1, 1);
            string downloadsPath = GetPath(KnownFolder.Downloads);
            book.SaveToFile(downloadsPath + "\\" + numero_factura + "_calidad.xls", ExcelVersion.Version97to2003);

            //actualizamos estado
            SqlCommand comm = null;
            comm = conprod.CreateCommand();
            comm.CommandText = ("UPDATE T SET T.ESTADO_CALIDAD = 1, " +
                                "T.FECHA_ACTUALIZACION_CAL = GETDATE() " +
                                "FROM omarsa..TBL_FACTURAS_CONTROL T  WHERE T.NUMERO_FACTURA ='" + numero_factura + "'");
            comm.CommandType = CommandType.Text;
            comm.CommandTimeout = 30;   //30 seconds
            comm.ExecuteNonQuery();
            

            ////// CONSULTAR DE CERTIFICADO SANITARIO 

            dt4 = new DataTable();
            adapt4 = new SqlDataAdapter("use omarsa  " +
                                        "IF OBJECT_ID('tempdb..#TMP_ECUAPASS') IS NOT NULL   BEGIN   DROP TABLE #TMP_ECUAPASS END "+
                                        "IF OBJECT_ID('tempdb..#Sanitario') IS NOT NULL      BEGIN   DROP TABLE #Sanitario END " +
                                        "CREATE TABLE #TMP_ECUAPASS  " +
                                        "([hc][varchar](max),[prdt_nm][varchar](max), idprod uniqueidentifier,[prdt_stn][varchar](max),[pck_qt][decimal](18, 6),[pck_ut][varchar](max),[prdt_nwt][decimal](18, 6),[nwt_ut][varchar](max),[pdtn_de][varchar](max),[metrica][varchar](max),[lote][varchar](max),[lot_no][varchar](max),[tipo][varchar](max)) " +
                                        "INSERT INTO #TMP_ECUAPASS " +
                                        "EXEC [dbo].[SP_CONTROL_CODIGOS_EMBARQUE_REAL_ECUAPASS_CHINA] '" + id_emb + "'  " +
                                        "CREATE TABLE #Sanitario([Subpartida Arancelaria] varchar(max),[Nombre de Producto] varchar(max), [Nombre Científico de Producto] varchar (max),[Cantidad de Embalaje de Producto] varchar(max),[Cantidad de Embalaje de Producto1] varchar(max),[Peso Neto de Producto]  varchar(max),[Peso Neto de Producto1]  varchar(max),[Número de Lote] varchar(max),[Fecha de Producción] varchar (max)  ) " +
                                        "INSERT INTO #Sanitario	VALUES ('hc','prdt_nm','prdt_stn','pck_qt','pck_ut','prdt_nwt','nwt_ut','lot_no','pdtn_de') " +
                                        "INSERT INTO #Sanitario	VALUES ('-','-','-','-','COM_0029','-','COM_0029','-','-') " +
                                        "INSERT INTO #Sanitario " +
                                        "SELECT " +
                                        "E.hc, " +
                                        "U.DVINGLES AS descripcion, " +
                                        "E.prdt_stn, " +
                                        "CAST(SUM(E.pck_qt)AS INT) pck_qt, " +
                                        "pck_ut, " +
                                        "CAST(SUM(prdt_nwt) AS INT)prdt_nwt, " +
                                        "nwt_ut, " +
                                        "e.lot_no, " +
                                        "MIN(E.pdtn_de) pdtn_de " +
                                        "FROM #TMP_ECUAPASS E " +
                                        "LEFT JOIN PRODUCTO P    ON E.IDPROD = P.ID " +
                                        "LEFT JOIN UD_PRODUCTO U ON U.ID = P.BOEXTENSION_ID  " +
                                        "group by e.hc,prdt_nm,u.DVINGLES,prdt_stn,pck_ut,nwt_ut,e.lot_no " +
                                        "SELECT* FROM #Sanitario "
                                        , conprod);
                                         adapt4.SelectCommand.CommandTimeout = 0;
                                        adapt4.Fill(dt4);
                                        dataGridView2.DataSource = dt4;

            //Descarga archivo Sanitario
            Workbook book1 = new Workbook();
            Worksheet sheet1 = book.Worksheets[0];
            book.Worksheets[0].InsertDataTable(this.dataGridView2.DataSource as DataTable, true, 1, 1);
            string downloadsPath1 = GetPath(KnownFolder.Downloads);
            book.SaveToFile(downloadsPath1 + "\\"+numero_factura+ "_sanitario.xls", ExcelVersion.Version97to2003);
            

            SqlCommand comm1 = null;
            comm1 = conprod.CreateCommand();
            comm1.CommandText = ("UPDATE T SET	T.ESTADO_SANITARIO = 1, " +
                                 "T.FECHA_ACTUALIZACION_SAN = GETDATE() " +
                                 "FROM omarsa..TBL_FACTURAS_CONTROL T  WHERE T.NUMERO_FACTURA ='" + numero_factura + "'");
            comm1.CommandType = CommandType.Text;
            comm1.CommandTimeout = 30;   //30 seconds
            comm1.ExecuteNonQuery();


        }

        public void button3_Click(object sender, EventArgs e)
        {
            //export specific data to Excel
            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            book.Worksheets[0].InsertDataTable(this.dataGridView1.DataSource as DataTable, true, 1, 1);
            //string downloadsPath = GetPath("\\Fileserver1\doc.export\Documentos_Calidad_Sanitarios");
            book.SaveToFile("\\Fileserver1\\doc.export\\Documentos_Calidad_Sanitarios" + "\\sample.xlsx", ExcelVersion.Version97to2003);
            //System.Diagnostics.Process.Start(downloadsPath+"\\sample.xls");

        }

        private void ecuapass_Tick(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            button1_Click(sender, e);
            button2_Click(sender, e);
            timer1.Start();



            /*
             
                       tmrEjecuta.Stop();
            Token1("Pjert93Ljs07l4wJfI92T9_ozGHX0ZMGP530AI3McGG2");
            tmrEjecuta.Start();
             */

        }
    }
}
