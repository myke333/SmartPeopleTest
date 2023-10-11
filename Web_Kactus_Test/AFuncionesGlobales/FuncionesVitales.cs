using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UITesting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Serialization;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
//using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.Configuration;
using OpenPop.Mime;
using OpenPop.Pop3;
using APITest;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;

namespace Web_Kactus_Test
{
    public class FuncionesVitales 
    {
        public static void ReporteErrores(List<string> Error, string testname)
        {
            string ruta = string.Format(@"C:\Reportes\Errores\{0}.txt", testname);
            StreamWriter sw = new StreamWriter(ruta);
            for (int i = 0; i < (Error.ToArray()).Length; i++)
            {
                sw.WriteLine(Error[i]);
            }
            sw.Close();
        }

        public static string Hora()
        {
            DateTime dateAndTime = DateTime.Now;
            string fecha = dateAndTime.ToString("ddMMyyyy_HHmmss");
            return fecha;
        }

        public static string CrearDocumentoWordDinamico(string app, string database, string modulo)
        {
            string NumCompilacion = ReadNumCompiler(app);
            string path = string.Format(@"C:\Reportes\Reportes{0}\Reportes_{1}\Ejecucion{2}", app, database, NumCompilacion);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string file = string.Format(@"{0}\Reportes_{1}_{2}.docx", path, modulo, Hora());
            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(file, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Prueba de ejecución realizada en el módulo: " + modulo));
            }
            return file;
        }

        public static string ReadNumCompiler(string app)
        {
            string PathXML = "";
            if (app == "SelfService")
            {
                PathXML = @"\\dwtfskscm\TfsScripts\TfsCompilerNumber\VariablesSelfService.xml";
            }
            else if (app == "Reclutamiento")
            {
                PathXML = @"\\dwtfskscm\TfsScripts\TfsCompilerNumber\VariablesReclutamiento.xml";
            }
            XmlDocument xmlDoc = new XmlDocument();

            xmlDoc.Load(PathXML);
            string NumCambios = "";
            foreach (XmlNode nodeProject in xmlDoc.GetElementsByTagName("VARIABLES"))
            {
                foreach (XmlNode nodePropertyGroup in nodeProject.ChildNodes)
                {
                    foreach (XmlNode nodeVerInfo_Keys in nodePropertyGroup.ChildNodes)
                    {
                        if (nodeVerInfo_Keys.Name == "GRUPOCAMBIOS")
                        {
                            NumCambios = nodeVerInfo_Keys.InnerText;
                        }
                    }

                }
            }
            return NumCambios;
        }
        
        public void Screenshot(string maestro, bool bandera, string file)
        {
            {
                Image MyImage = null;
                string UrlImage = @"C:\Reportes\" + maestro + ".bmp";
                MyImage = UITestControl.Desktop.CaptureImage();
                MyImage.Save(UrlImage, System.Drawing.Imaging.ImageFormat.Bmp);
                InsertAPicture(file, UrlImage, maestro, bandera);

            }
        }

        public static void InsertAPicture(string document, string UrlImage, string maestro, bool bandera)
        {
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(UrlImage, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), maestro, bandera, UrlImage);
            }
        }

        

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, string maestro, bool bandera, string UrlImage)
        {
            int iWidth = 0;
            int iHeight = 0;
            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(UrlImage))
            {
                iWidth = bmp.Width;
                iHeight = bmp.Height;
            }
            iWidth = (int)Math.Round((decimal)iWidth * 4000);
            iHeight = (int)Math.Round((decimal)iHeight * 4000);

            var element =

                new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = iWidth, Cy = iHeight },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.bmp"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.None
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 10000L, Y = 10000L },
                                             new A.Extents() { Cx = iWidth, Cy = iHeight }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            if (bandera)
            {
                wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(maestro))));
            }
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

        }

        public void ConverWordToPDF(string RutaArchivo, string database)
        {
            string[] split = RutaArchivo.Split('\\');
            string Ruta = "";
            for (int i = 0; i < (split.Length - 1); i++)
            {
                Ruta = Ruta + split[i] + @"\";
            }
            Ruta = string.Format(Ruta + @"ArchivosPDF_{0}\", database);
            string Archivo = split[split.Length - 1];
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;
            word.Visible = false;
            word.ScreenUpdating = false;
            Object FileName = (Object)RutaArchivo;
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref FileName, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();
            if (!Directory.Exists(Ruta))
            {
                Directory.CreateDirectory(Ruta);
            }
            object Filter = Archivo.Replace(".docx", ".pdf");
            object outputFileName = Ruta + Filter;
            object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            doc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            ((Microsoft.Office.Interop.Word.Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
            doc = null;
            ((Microsoft.Office.Interop.Word.Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
            Process[] processes2 = Process.GetProcessesByName("WINWORD");
            if (processes2.Length > 0)
            {
                for (int i = 0; i < processes2.Length; i++)
                {
                    processes2[i].Kill();
                }
            }
        }

        static public void LimpiarProcesos()
        {
            Process[] processes = Process.GetProcessesByName("AcroRd32");
            if (processes.Length > 0)
            {
                for (int i = 0; i < processes.Length; i++)
                {
                    processes[i].Kill();
                }
            }
            Process[] processes1 = Process.GetProcessesByName("EXCEL");
            if (processes1.Length > 0)
            {
                for (int i = 0; i < processes1.Length; i++)
                {
                    processes1[i].Kill();
                }
            }
            Process[] processes2 = Process.GetProcessesByName("WINWORD");
            if (processes2.Length > 0)
            {
                for (int i = 0; i < processes2.Length; i++)
                {
                    processes2[i].Kill();
                }
            }

            Process[] processes3 = Process.GetProcessesByName("notepad");
            if (processes3.Length > 0)
            {
                for (int i = 0; i < processes3.Length; i++)
                {
                    processes3[i].Kill();
                }
            }
        }

        public DataSet ConsultarSqlOra(string Sentencia, string user, string database)
        {
            DataSet dataSet = new DataSet();
            DataTable dTable = new DataTable("dTable");
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;

            if (database == "SQL")
            {
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dataSet;
            }
            else if (database == "ORA")
            {
                OracleConnection cnn = new OracleConnection(ConnectionString2);
                cnn.Open();
                OracleDataAdapter adapter = new OracleDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dataSet;
            }
            return null;
        }

        static public void UpdateDeleteInsert(string Sentencia, string Database, string User)
        {
            string ConnectionString = ConfigurationManager.ConnectionStrings[User].ConnectionString;
            switch (Database.ToUpper())
            {
                case "SQL":

                    SqlConnection sqlConnection = new SqlConnection(ConnectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    SqlTransaction sqlTransaction;
                    sqlTransaction = sqlConnection.BeginTransaction();
                    sqlCommand.Connection = sqlConnection;
                    sqlCommand.Transaction = sqlTransaction;
                    try
                    {
                        sqlCommand.CommandText = Sentencia;
                        sqlCommand.ExecuteNonQuery();
                        sqlTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        sqlTransaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    sqlConnection.Close();
                    break;

                case "ORA":

                    OracleConnection oracleConnection = new OracleConnection(ConnectionString);
                    oracleConnection.Open();
                    OracleCommand oracleCommand = oracleConnection.CreateCommand();
                    OracleTransaction oracleTransaction;
                    oracleTransaction = oracleConnection.BeginTransaction();
                    oracleCommand.Connection = oracleConnection;
                    oracleCommand.Transaction = oracleTransaction;
                    try
                    {
                        oracleCommand.CommandText = Sentencia;
                        oracleCommand.ExecuteNonQuery();
                        oracleTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        oracleTransaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    oracleConnection.Close();
                    break;
                default:
                    break;
            }



        }

        static public void ValidarCorreo(
            out string Remitente,
            out string DescRemitente,
            out string Asunto,
            out string Body,
            string Correo, 
            string Password)
        {
            string UserEmail = $"recent:{Correo}";
            int port = 995;
            bool UseSSL = true;
            string Hostname = "pop.gmail.com";
            Remitente = string.Empty;
            DescRemitente = string.Empty;
            Asunto = string.Empty;
            Body = string.Empty;
            StringBuilder builder = new StringBuilder();

            Pop3Client popClient = new Pop3Client();
            popClient.Connect(Hostname, port, UseSSL);
            popClient.Authenticate(UserEmail, Password);
            var count = popClient.GetMessageCount();

            Message message = popClient.GetMessage(count);
            Remitente = message.Headers.From.Address;
            DescRemitente = message.Headers.From.DisplayName;
            Asunto = message.Headers.Subject;
            MessagePart plainText = message.FindFirstPlainTextVersion();
            builder.Append(plainText.GetBodyAsText());
            Body = builder.ToString();
            Thread.Sleep(500);
            popClient.DeleteMessage(count);
            popClient.Disconnect();
        }
    }
}