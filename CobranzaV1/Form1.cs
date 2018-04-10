using ClosedXML.Excel;
using System;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;
using System.Drawing;
using System.IO;
using System.Xml;
using NPOI.XSSF.UserModel;

namespace CobranzaV1
{
    public partial class Form1 : Form
    {
        int ANOSTOTAL = 9;
        int MUNICIPIOS = 348;
        public Form1()
        {          
            InitializeComponent();      
        }

        private void ConsultaWebCLASIFICARPRESUPUESTARIO()
        {
            string ANO = textBoxANO.Text;
            string ruta = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario";
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
                //Directory.CreateDirectory(ruta + "\\2008");
                //Directory.CreateDirectory(ruta + "\\2009");
                //Directory.CreateDirectory(ruta + "\\2010");
                //Directory.CreateDirectory(ruta + "\\2011");
                //Directory.CreateDirectory(ruta + "\\2012");
                //Directory.CreateDirectory(ruta + "\\2013");
                //Directory.CreateDirectory(ruta + "\\2014");
                //Directory.CreateDirectory(ruta + "\\2015");
                //Directory.CreateDirectory(ruta + "\\2016");
                //Directory.CreateDirectory(ruta + "-xlsx\\2008");
                //Directory.CreateDirectory(ruta + "-xlsx\\2009");
                //Directory.CreateDirectory(ruta + "-xlsx\\2010");
                //Directory.CreateDirectory(ruta + "-xlsx\\2011");
                //Directory.CreateDirectory(ruta + "-xlsx\\2012");
                //Directory.CreateDirectory(ruta + "-xlsx\\2013");
                //Directory.CreateDirectory(ruta + "-xlsx\\2014");
                //Directory.CreateDirectory(ruta + "-xlsx\\2015");
                //Directory.CreateDirectory(ruta + "-xlsx\\2016");
            }

            if (!Directory.Exists(ruta + "\\"+ ANO ))
            {
                Directory.CreateDirectory(ruta + "\\"+ ANO);
                Directory.CreateDirectory(ruta + "-xlsx\\"+ ANO);
            }

            string anocorto = ANO.Substring(2, 2);
            int numero = Convert.ToInt32(anocorto);
            numero++;
            

            int AAA = 183;
            
            //--------AÑO 2016--------
            AAA = 183;
            while (AAA < 529)
            {
                string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]="+ numero + "&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
                using (WebClient wc = new WebClient())
                {
                    wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                    string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\"+ ANO +"\\"+ ANO +"_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
                    wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
                    ConvertirAExcel(URLLOCAL + AAA + ".xls");
                    AAA++;
                    label5.Text = "MUNICIPIOS DESCARGADOS" + (AAA - 183).ToString();
                }
            }
            #region clasificador presupuestario
            ////--------AÑO 2015--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=16&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2015\\2015_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2014--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=15&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2014\\2014_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2013--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=14&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2013\\2013_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2012--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=13&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2012\\2012_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2011--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=12&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2011\\2011_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2010--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=11&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2010\\2010_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2009--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=10&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2009\\2009_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2008--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/clasificador_presupuestario/obtener_clasificador_presupuestario.php?area[]=T&subarea[]=T&areagestion[]=null&agrupador[]=T&cuentas[]=T&periodos[]=9&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\ClasificadorPresupuestario\\2008\\2008_Clasificador_Presupuestario_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            #endregion
        }

        private void ConsultaWebDATOSMUNICIPALES()
        {
            string ANO = textBoxANO.Text;
            string ruta = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales";
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
                //Directory.CreateDirectory(ruta + "\\2008");
                //Directory.CreateDirectory(ruta + "\\2009");
                //Directory.CreateDirectory(ruta + "\\2010");
                //Directory.CreateDirectory(ruta + "\\2011");
                //Directory.CreateDirectory(ruta + "\\2012");
                //Directory.CreateDirectory(ruta + "\\2013");
                //Directory.CreateDirectory(ruta + "\\2014");
                //Directory.CreateDirectory(ruta + "\\2015");
                //Directory.CreateDirectory(ruta + "\\2016");
                //Directory.CreateDirectory(ruta + "-xlsx\\2008");
                //Directory.CreateDirectory(ruta + "-xlsx\\2009");
                //Directory.CreateDirectory(ruta + "-xlsx\\2010");
                //Directory.CreateDirectory(ruta + "-xlsx\\2011");
                //Directory.CreateDirectory(ruta + "-xlsx\\2012");
                //Directory.CreateDirectory(ruta + "-xlsx\\2013");
                //Directory.CreateDirectory(ruta + "-xlsx\\2014");
                //Directory.CreateDirectory(ruta + "-xlsx\\2015");
                //Directory.CreateDirectory(ruta + "-xlsx\\2016");
            }

            if (!Directory.Exists(ruta + "\\" + ANO))
            {
                Directory.CreateDirectory(ruta + "\\" + ANO);
                Directory.CreateDirectory(ruta + "-xlsx\\" + ANO);
            }

            string anocorto = ANO.Substring(2, 2);
            int numero = Convert.ToInt32(anocorto);
            numero++;

            
            int AAA = 183;
            //--------AÑO 2016--------
            AAA = 183;
            while (AAA < 529)
            {
                string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]="+ numero +"& regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
                using (WebClient wc = new WebClient())
                {
                    wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                    string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\" + ANO + "\\" + ANO + "_Datos_Municipales_Sin-Corrección-Monetaria";
                    wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
                    ConvertirAExcel(URLLOCAL + AAA + ".xls");
                    AAA++;
                    label5.Text = "MUNICIPIOS DESCARGADOS" + (AAA - 183).ToString();
                }
            }
            #region AÑOS DESCARGA
            ////--------AÑO 2015--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=16&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2015\\2015_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }   
            //}
            ////--------AÑO 2014--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=15&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2014\\2014_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2013--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=14&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2013\\2013_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2012--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=13&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2012\\2012_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2011--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=12&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2011\\2011_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2010--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=11&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2010\\2010_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2009--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=10&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2009\\2009_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}
            ////--------AÑO 2008--------
            //AAA = 183;
            //while (AAA < 529)
            //{
            //    string URI = "http://datos.sinim.gov.cl/datos_municipales/obtener_datos_municipales.php?area[]=T&subarea[]=T&variables[]=T&periodos[]=9&regiones[]=T&municipios[]=" + AAA + "&corrmon=false";
            //    using (WebClient wc = new WebClient())
            //    {
            //        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            //        string URLLOCAL = "C:\\Users\\Public\\DESCARGAAUTOMATICA\\DatosMunicipales\\2008\\2008_Datos_Municipales_Sin-Corrección-Monetaria";
            //        wc.DownloadFile(URI, URLLOCAL + AAA + ".xls");
            //        ConvertirAExcel(URLLOCAL + AAA + ".xls");
            //        AAA++;
            //    }
            //}

            #endregion
        }

        public void ConvertirAExcel(string direccion)
        {
            StreamReader lector = File.OpenText(direccion);
            string linea;
            string texto="";
            do
            {
                linea = lector.ReadLine();
                if (linea != null)
                {
                    texto=texto+ linea;
                }
            }
            while (linea != null);
            lector.Close();

            XmlDocument xm = new XmlDocument();
            xm.LoadXml(texto);
            xm.Save(direccion);
            xm = null;

            //DirectoryInfo di = new DirectoryInfo(@"C:\Users\Luis León\Desktop\CLASIFICADOR PRESUPUESTARIO\20156666");
            XmlDocument xDoc = new XmlDocument();
            using (FileStream fs = new FileStream(direccion, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xDoc.Load(fs);
            }
            XmlNodeList xPersonas = ((XmlElement)((XmlElement)xDoc.GetElementsByTagName("Worksheet")[0]).GetElementsByTagName("Table")[0]).GetElementsByTagName("Row");
            xDoc = null;

            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Hoja1");
            // Add header labels
            int filas = xPersonas.Count;
            var rowIndex = 0;

            for (int i = 0; i < filas; i++)
            {
                XmlNodeList xCell = ((XmlElement)xPersonas[i]).GetElementsByTagName("Cell");
                int contadorCol = xCell.Count;
                var row = sheet.CreateRow(rowIndex);
                for (int j = 0; j < contadorCol; j++)
                {
                    row.CreateCell(j).SetCellValue(xCell[j].InnerText);
                }
                rowIndex++;
            }

            // Save the Excel spreadsheet to a file on the web server's file system
            direccion = direccion.Replace("DatosMunicipales", "DatosMunicipales-xlsx");
            direccion = direccion.Replace("ClasificadorPresupuestario", "ClasificadorPresupuestario-xlsx");
            direccion = direccion.Replace(".xls", ".xlsx");
            using (var fileData = new FileStream(direccion, FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ConsultaWebCLASIFICARPRESUPUESTARIO();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConsultaWebDATOSMUNICIPALES();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
