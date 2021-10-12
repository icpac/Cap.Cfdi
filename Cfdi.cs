#region tit
/*
{+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++)
{                                                                   }
{     tit Cap control administrativo personal                       }
{     Cuestiones sobre cfdi                                         }
{                                                                   }
{*******************************************************************}
 */
#endregion

using System;
using System.Diagnostics;
using System.IO;
using Cap.Fe.BusinessObjects;
using Cap.Generales.BusinessObjects.Empresa;
using Cap.Generales.BusinessObjects.Unidades;
using Cap.Inventarios.BusinessObjects;
using Cap.Nomina.BusinessObjects;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.Spreadsheet;

namespace Cap.Cfdi
{
    public class Cfdi
    {
        // Ahi graba los pdfs y de ahi lo leemos.
        /// <summary>
        /// Crea la ruta de acuerda a la base y a la fecha del documento
        /// </summary>
        /// <param name="ruta"></param>
        /// <param name="fecha"></param>
        /// <param name="tipo"></param>
        public static string CreaDirs(string ruta, DateTime fecha, string tipo)
        {
            string aux = Path.Combine(ruta, fecha.Year.ToString());

            if (!Directory.Exists(aux))
                Directory.CreateDirectory(aux);

            aux = Path.Combine(aux, fecha.Month.ToString());

            if (!Directory.Exists(aux))
                Directory.CreateDirectory(aux);

            aux = Path.Combine(aux, fecha.Day.ToString());

            if (!Directory.Exists(aux))
                Directory.CreateDirectory(aux);

            return aux;
        }


        /*
        Para crear llave publica
        C:\Program Files\GnuWin32\bin>openssl.exe pkcs8 -inform DER -in C:\aaa010101aaa_CSD_01.key -passin pass:a0123456789 -out C:\aaa010101aaa_CSD_01.key.pem

        Para crear llave privada
        C:\Program Files\GnuWin32\bin>openssl.exe x509 -inform DER -outform PEM -in C:\aaa010101aaa_CSD_01.cer -pubkey -out C\aaa010101aaa_CSD_01.cer.pem
        */
        #region + Crea Archivos Pem
        public static void CreaArchivosPEM(string filc, string filk, string pass)
        {
            if (!string.IsNullOrEmpty(filk)
                && !string.IsNullOrEmpty(pass)
                && !string.IsNullOrEmpty(filc))
            {
                string ArchKeyPem = filc + ".pem";

                if (!File.Exists(ArchKeyPem))
                {
                    Process myproc;

                    try
                    {
                        myproc = new Process();

                        myproc.StartInfo.UseShellExecute = false;
                        myproc.EnableRaisingEvents = false;
                        ProcessStartInfo psi = new ProcessStartInfo(
                            "openssl.exe", " pkcs8 -inform DER -in \"" +
                            filk + "\" -passin pass:" +
                            pass + " -out \"" + filk + ".pem\"");
                        myproc.StartInfo = psi;
                        myproc.Start();

                        ProcessStartInfo psi2 = new ProcessStartInfo(
                            "openssl.exe", " x509 -inform DER -outform PEM -in \""
                            + filc + "\" -pubkey -out \"" + filc + ".pem\"");
                        myproc.StartInfo = psi2;
                        myproc.Start();
                    }
                    catch (Exception ee)
                    {
                        throw new Exception(string.Format("No se pudieron crear archivos PEM\n {0}", ee.Message));
                    }
                }
            }
        }
        #endregion


        // Nómina
        // c_Banco, c_OrigenRecurso, c_TipoHoras, c_TipoOtroPago, 

        public static void LoadData(FilterCFDI obj, IObjectSpace obs, string cnnS)
        {
            Workbook wb = new Workbook();
            DocumentFormat tipoLibro = Path.GetExtension(obj.File.FileName.ToUpper()).Contains("XLSX") ? DocumentFormat.Xlsx : DocumentFormat.Xls;
            wb.LoadDocument(obj.File.Content, tipoLibro);

            if (wb != null && wb.Worksheets.Count > 0)
            {
                if (obj.FrmnsPg)
                {
                    Worksheet fp = wb.Worksheets["c_FormaPago"];

                    cargaFormaPago(fp, obj.Rngln, obs);
                }

                if (obj.MtdPg)
                {
                    Worksheet mtd = wb.Worksheets["c_MetodoPago"];

                    cargaMetodoPago(mtd, obj.Rngln, obj.RnglnFnl, obs);
                }

                if (obj.Rgmn)
                {
                    Worksheet mtd = wb.Worksheets["c_RegimenFiscal"];

                    cargaRegimen(mtd, obj.Rngln, obs);
                }

                if (obj.UsCFDI)
                {
                    Worksheet mtd = wb.Worksheets["c_USOCFDI"];

                    cargaUsoCfdi(mtd, 7, obs);
                }

                if (obj.PrdctsSrvcs)
                {
                    Worksheet mtd = wb.Worksheets["c_ClaveProdServ"];

                    cargaProdServ(mtd, obj.Rngln, obj.RnglnFnl, obs, cnnS);
                }

                if (obj.Undds)
                {
                    Worksheet mtd = wb.Worksheets["c_ClaveUnidad"];

                    cargaUndds(mtd, obj.Rngln, obs);
                }

                if (obj.Mnds)
                {
                    Worksheet md = wb.Worksheets["c_Moneda"];

                    cargaMonedas(md, 6, obs);
                }

                if (obj.Impsts)
                {
                    Worksheet md = wb.Worksheets["c_Impuesto"];

                    cargaImpuestos(md, 6, obs);
                }

                if (obj.TpRlcn)
                {
                    Worksheet md = wb.Worksheets["c_TipoRelacion"];

                    cargaTipoRelacion(md, 6, obs);
                }

                if (obj.TpPrcpcns)
                {
                    Worksheet md = wb.Worksheets["c_TipoPercepcion"];

                    cargaTipoPercepcion(md, obj.Rngln, obs);
                }

                if (obj.TpDdcns)
                {
                    Worksheet md = wb.Worksheets["c_TipoDeduccion"];

                    cargaTipoDeduccion(md, obj.Rngln, obs);
                }

                if (obj.Rsg)
                {
                    Worksheet md = wb.Worksheets["c_RiesgoPuesto"];

                    cargaRiesgos(md, obj.Rngln, obs);
                }

                if (obj.Incpcdd)
                {
                    Worksheet md = wb.Worksheets["c_TipoIncapacidad"];

                    cargaIncapacidad(md, obj.Rngln, obs);
                }

                if (obj.TpCntrt)
                {
                    Worksheet md = wb.Worksheets["c_TipoContrato"];

                    cargaTipoContratos(md, obj.Rngln, obs);
                }

                if (obj.Prdcdd)
                {
                    Worksheet md = wb.Worksheets["c_PeriodicidadPago"];

                    cargaCatalogosNomina(md, obj.Rngln, 0, obs);
                }

                if (obj.TpNmn)
                {
                    Worksheet md = wb.Worksheets["c_TipoNomina"];

                    cargaCatalogosNomina(md, obj.Rngln, 1, obs);
                }

                if (obj.TpRgmnCntrt)
                {
                    Worksheet md = wb.Worksheets["c_TipoRegimen"];

                    cargaTipoRegimenContratos(md, obj.Rngln, obs);
                }

                if (obj.TpJrnd)
                {
                    Worksheet md = wb.Worksheets["c_TipoJornada"];

                    cargaTipoJornada(md, obj.Rngln, obs);
                }
            }
        }

        private static void cargaFormaPago(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i, clv;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    clv = CeldaEntero(ws, i, iclave);

                    if (clv > 0)
                    {
                        if (objectSpace.FindObject<Pago>(new BinaryOperator("Clv", clv.ToString().PadLeft(2, '0'))) == null)
                        {
                            Pago pg = objectSpace.CreateObject<Pago>();

                            pg.Clv = clv.ToString().PadLeft(2, '0');
                            pg.Descrip = CeldaTexto(ws, i, idscrpcn);
                            pg.Tipo = TipoPago.FormaPago;
                        }
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaMetodoPago(Worksheet ws, int renIni, 
            int renFin, IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 2)
            {
                int i;
                string clv;
                IObjectSpace objectSpace = obs; 
                int rf = renFin == 0 ? rango.RowCount : renFin;

                for (i = renIni - 1; i < rf; i++)
                {
                    clv = CeldaTexto(ws, i, iclave);

                    if (!string.IsNullOrEmpty(clv))
                    {
                        if (objectSpace.FindObject<Pago>(new BinaryOperator("Clv", clv)) == null)
                        {
                            Pago pg = objectSpace.CreateObject<Pago>();

                            pg.Clv = clv;
                            pg.Descrip = CeldaTexto(ws, i, idscrpcn);
                            pg.Tipo = TipoPago.Metodo;
                        }
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        /*Ago 2020 Están cambiando el tamaño de los campos
         y aunque hay que verificarlo, una opción es que no se 
        caiga por ello */
        private static void cargaRegimen(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 2)
            {
                int i;
                int clv;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    clv = CeldaEntero(ws, i, iclave);

                    if (clv > 0)
                    {
                        if (objectSpace.FindObject<Regimen>(new BinaryOperator("Clv", clv.ToString())) == null)
                        {
                            Regimen rgm = objectSpace.CreateObject<Regimen>();

                            rgm.Clv = clv.ToString();
                            rgm.Dscrpcn = CeldaTexto(ws, i, idscrpcn);
                        }
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaUsoCfdi(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 2)
            {
                int i;
                string clv;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    clv = CeldaTexto(ws, i, iclave);

                    if (!string.IsNullOrEmpty(clv))
                    {
                        if (objectSpace.FindObject<Pago>(new BinaryOperator("Clv", clv)) == null)
                        {
                            Pago pg = objectSpace.CreateObject<Pago>();

                            pg.Clv = clv;
                            pg.Descrip = CeldaTexto(ws, i, idscrpcn);
                            pg.Tipo = TipoPago.UsoCFDI;
                        }
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaProdServ(Worksheet ws, int renIni, 
            int renFin, IObjectSpace obs, string cnS)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 2)
            {
                int i;
                string clv;
                IObjectSpace objectSpace = obs;
                int rf = renFin == 0 ? rango.RowCount : renFin;

                for (i = renIni - 1; i < rf; i++)
                {
                    CellRange celda = ws.Cells[i, iclave];
                    if (celda.Value.Type == CellValueType.Text)
                        clv = CeldaTexto(ws, i, iclave);
                    else
                    {
                        int cv = CeldaEntero(ws, i, iclave);
                        clv = cv.ToString();
                    }


                    if (!string.IsNullOrEmpty(clv))
                    {
                        if (objectSpace.FindObject<ProductoServicio>(new BinaryOperator("Clv", clv)) == null)
                        {
                            string insertQuery = string.Empty;
                            if (!cnS.Contains("postgres"))
                            {
                                insertQuery = string.Format(@"INSERT INTO [ProductoServicio] ([Oid], [Clv], [Dscrpcn], [Tp]) VALUES ('{0}', '{1}', @d, {2})", 
                                    Guid.NewGuid(), clv, 
                                    clv.CompareTo("70101500") < 0
                                    ? ETipoProductoServicioFe.Producto.GetHashCode()
                                    : ETipoProductoServicioFe.Servicio.GetHashCode());
                            }
                            else
                            {
                                insertQuery = string.Format("INSERT INTO \"ProductoServicio\" (\"Oid\", \"Clv\", \"Dscrpcn\", \"Tp\") VALUES ('{0}', '{1}', @d, {2})",
                                    Guid.NewGuid(), clv, clv.CompareTo("70101500") < 0
                                    ? ETipoProductoServicioFe.Producto
                                    : ETipoProductoServicioFe.Servicio);  
                            }
                            ((DevExpress.ExpressApp.Xpo.XPObjectSpace)objectSpace).Session.ExecuteNonQuery(
                                insertQuery, new string[] { "d" }, new object[] { CeldaTexto(ws, i, idscrpcn) });
                        }
                    }
                }
            }
        }

        private static void cargaUndds(Worksheet ws, int renIni, IObjectSpace obs)
        {
            int iclave, inmbr, idscrpcn, isgls;
            string colClave = "A";
            string colNmbr = "B";
            string colDscrpcn = "C";
            string colSgls = "F";

            iclave = IndiceDeColumna(ws, colClave);
            inmbr = IndiceDeColumna(ws, colNmbr);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);
            isgls = IndiceDeColumna(ws, colSgls);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 2)
            {
                int i;
                string clv;
                IObjectSpace objectSpace = obs; 

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    CellRange celda = ws.Cells[i, iclave];
                    if (celda.Value.Type == CellValueType.Text)
                        clv = CeldaTexto(ws, i, iclave);
                    else
                    {
                        int cv = CeldaEntero(ws, i, iclave);
                        clv = cv.ToString();
                    }


                    if (!string.IsNullOrEmpty(clv))
                    {
                        if (objectSpace.FindObject<Unidad>(new BinaryOperator("Clave", clv)) == null)
                        {

                            Unidad und = objectSpace.CreateObject<Unidad>();

                            und.Clave = clv;
                            und.Nmbr = CeldaTexto(ws, i, inmbr);
                            und.Descripcion = CeldaTexto(ws, i, idscrpcn);
                            und.Siglas = CeldaTexto(ws, i, isgls);


                            if (i % 64 == 0)
                            {
                                objectSpace.CommitChanges();
                            }
                        }
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaMonedas(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs; 

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    string aux = CeldaTexto(ws, i, iclave);

                    if (!string.IsNullOrEmpty(aux) 
                        && objectSpace.FindObject<Moneda>(new BinaryOperator("Clave", Moneda.ClaveFto(aux))) == null)
                    {
                        Moneda mnd = objectSpace.CreateObject<Moneda>();

                        mnd.Clave = aux;
                        mnd.Descripcion = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaImpuestos(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn, irtn, itrs;
            string colClave = "A";
            string colDscrpcn = "B";
            string colRtnc = "C";
            string colTrsld = "D";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);
            irtn = IndiceDeColumna(ws, colRtnc);
            itrs = IndiceDeColumna(ws, colTrsld);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = aux.ToString();
                    string aux3 = aux2.PadLeft(3, '0');

                    if (!string.IsNullOrEmpty(aux3) 
                        && objectSpace.FindObject<Impuesto>(new BinaryOperator("Clv", aux3)) == null)
                    {
                        Impuesto impt = objectSpace.CreateObject<Impuesto>();

                        impt.Clv = aux3;
                        impt.Dscrpcn = CeldaTexto(ws, i, idscrpcn);
                        impt.Rtncn = CeldaTexto(ws, i, irtn) == "Si";
                        impt.Trsld = CeldaTexto(ws, i, itrs) == "Si";
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaTipoRelacion(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(2, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    GroupOperator fil = new GroupOperator();

                    fil.Operands.Add(new BinaryOperator("Clv", aux3));
                    fil.Operands.Add(new BinaryOperator("Tipo", TipoPago.TipoRelacion));

                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<Pago>(fil) == null)
                    {
                        Pago pg = objectSpace.CreateObject<Pago>();

                        pg.Clv = aux3;
                        pg.Descrip = CeldaTexto(ws, i, idscrpcn);
                        pg.Tipo = TipoPago.TipoRelacion;

                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaTipoPercepcion(Worksheet ws, 
            int renIni, IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(3, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<TipoPer>(new BinaryOperator("Clave", aux3)) == null)
                    {
                        TipoPer impt = objectSpace.CreateObject<TipoPer>();

                        impt.Clave = aux3;
                        impt.Descrip = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaTipoDeduccion(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(3, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<TipoDed>(new BinaryOperator("Clave", aux3)) == null)
                    {
                        TipoDed impt = objectSpace.CreateObject<TipoDed>();

                        impt.Clave = aux3;
                        impt.Descrip = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaRiesgos(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2;
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<Riesgo>(new BinaryOperator("Clave", aux3)) == null)
                    {
                        Riesgo impt = objectSpace.CreateObject<Riesgo>();

                        impt.Clave = aux3;
                        impt.Descrip = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaIncapacidad(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(2, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<Incapacidad>(new BinaryOperator("Clave", aux3)) == null)
                    {
                        Incapacidad impt = objectSpace.CreateObject<Incapacidad>();

                        impt.Clave = aux3;
                        impt.Descrip = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaTipoContratos(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs; 

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(2, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<Contratacion>(new BinaryOperator("Clave", aux3)) == null)
                    {
                        Contratacion impt = objectSpace.CreateObject<Contratacion>();

                        impt.Clave = aux3;
                        impt.Descrip = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaCatalogosNomina(Worksheet ws, int renIni, int tpCtlg, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(2, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    
                    //*TIT Oct 2018 dónde está CatalogoNomina?
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<CatalogoNomina>(new BinaryOperator("Clv", aux3)) == null)
                    {
                        CatalogoNomina impt = objectSpace.CreateObject<CatalogoNomina>();

                        impt.Clv = aux3;
                        impt.Dscrpcn = CeldaTexto(ws, i, idscrpcn);
                        if (tpCtlg == 0)
                            impt.Tp = ETIPONOMINA.Periodicidad;
                        else if (tpCtlg == 1)
                            impt.Tp = ETIPONOMINA.TipoNomina;
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaTipoRegimenContratos(Worksheet ws, int renIni, 
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(2, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<RegimenContrato>(new BinaryOperator("Clave", aux3)) == null)
                    {
                        RegimenContrato impt = objectSpace.CreateObject<RegimenContrato>();

                        impt.Clave = aux3;
                        impt.Descrip = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }

        private static void cargaTipoJornada(Worksheet ws, int renIni,
            IObjectSpace obs)
        {
            int iclave, idscrpcn;
            string colClave = "A";
            string colDscrpcn = "B";

            iclave = IndiceDeColumna(ws, colClave);
            idscrpcn = IndiceDeColumna(ws, colDscrpcn);

            CellRange rango = ws.GetUsedRange();
            if (rango.ColumnCount >= 3)
            {
                int i;
                IObjectSpace objectSpace = obs;

                for (i = renIni - 1; i < rango.RowCount; i++)
                {
                    int aux = CeldaEntero(ws, i, iclave);
                    string aux2 = string.Empty;
                    string aux3 = string.Empty;
                    if (aux != 0)
                    {
                        aux2 = aux.ToString();
                        aux3 = aux2.PadLeft(2, '0');
                    }
                    else
                    {
                        aux2 = CeldaTexto(ws, i, iclave);
                        aux3 = aux2.Trim();
                    }
                    if (!string.IsNullOrEmpty(aux3)
                        && objectSpace.FindObject<Jornada>(new BinaryOperator("Clave", aux3)) == null)
                    {
                        Jornada impt = objectSpace.CreateObject<Jornada>();

                        impt.Clave = aux3;
                        impt.Dscrpcn = CeldaTexto(ws, i, idscrpcn);
                        //  viene un dato de numero de decimales que no se si se use
                    }
                }
                objectSpace.CommitChanges();
            }
        }







        public static int IndiceDeColumna(Worksheet ws, string encabezado)
        {
            if (ws == null || string.IsNullOrEmpty(encabezado))
                return -1;
            return ws.Columns[encabezado].Index;
        }

        public static int CeldaEntero(Worksheet ws, int r, int c)
        {
            return Convert.ToInt32(CeldaNumero(ws, r, c));
        }

        public static string CeldaTexto(Worksheet ws, int r, int c)
        {
            string texto;
            CellRange celda = ws.Cells[r, c];
            if (celda.Value.Type == CellValueType.None || celda.Value.Type == CellValueType.Error)
                texto = "";
            else
                texto = celda.Value.TextValue.Trim();
            return texto;
        }

        public static double CeldaNumero(Worksheet ws, int r, int c)
        {
            double valor = 0.0;
            try
            {
                CellRange celda = ws.Cells[r, c];
                if (celda.Value.Type == CellValueType.Numeric)
                    valor = celda.Value.NumericValue;
                else if (celda.Value.Type == CellValueType.Text)
                    valor = Convert.ToDouble(celda.Value.ToString(), null);
            }
            catch
            {
                valor = 0.0;
            }
            return valor;
        }
    }
}
