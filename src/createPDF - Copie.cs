using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using Newtonsoft.Json;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
//using System.Windows.Documents;
using System.Windows.Media;
using System.Xml.Linq;
using VMS.OIS.ARIALocal.WebServices.Document.Contracts;
using VMS.TPS.Common.Model;
using VMS.TPS.Common.Model.API;
using MDOM = MigraDoc.DocumentObjectModel;

namespace catchDose
{
    internal class createPDF
    {
        private MDOM.Document migraDoc;
        public string pdfFilePath { get; set; }
        public string jsonFilePath { get; set; }
        PreliminaryInformation pinfo { get; set; }
        public ScriptContext ctx { get; set; }

        public string docteur { get; set; }
        public createPDF(ScriptContext _ctx, PreliminaryInformation _pinfo, string jsonPath) //constructor
        {
            pinfo = _pinfo;
            int ic = 0;
            ctx = _ctx;

            string jsonContent = "";
            jsonFilePath = jsonPath;

            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(jsonPath);
            string pdfDirectory = Path.GetDirectoryName(jsonPath).Replace(@"\json", @"\pdf");
            string dateString = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            pdfFilePath = Path.Combine(pdfDirectory, fileNameWithoutExt + "_" + dateString + ".pdf");
          
            if (File.Exists(jsonPath))
            {
                jsonContent = File.ReadAllText(jsonPath);
                uniqueJsonFileName ujfn = new uniqueJsonFileName();
                ujfn.unmakefilename(Path.GetFileName(jsonPath));
                string ipp = ujfn.IPP; // "12345"
                string course = ujfn.courseID; // "COURSE1"
                string plan = ujfn.planID; // "PLAN_A"
            }
            else
            {
                MessageBox.Show("Le fichier json n'existe pas:\n" + jsonPath);
                return;
            }


            migraDoc = new MDOM.Document();
            MDOM.Section section = migraDoc.AddSection();
            section.PageSetup.Orientation = MDOM.Orientation.Portrait;

            #region header

            string logoPath = Directory.GetCurrentDirectory() + @"\img\logo.jpg";
            if(!File.Exists(logoPath))
            {
                MessageBox.Show("Le logo est introuvable :\n" + logoPath);
            }
            else
            {
                MigraDoc.DocumentObjectModel.Shapes.Image image = section.AddImage(logoPath);
                image.Width = "5cm";
                image.LockAspectRatio = true;
                image.Left = MigraDoc.DocumentObjectModel.Shapes.ShapePosition.Center;

                 Paragraph spacer = section.AddParagraph();
                spacer.Format.SpaceBefore = "0.5cm";
                spacer.Format.SpaceAfter = "0.1cm";
                
            }

            MDOM.Paragraph title = section.AddParagraph("RAPPORT DE LIMITES DE DOSES");
            title.Format.Alignment = MDOM.ParagraphAlignment.Center;
            title.Format.Font.Size = 18;
            title.Format.Font.Bold = true;
            title.Format.Font.Color = MDOM.Colors.DarkBlue;
            title.Format.SpaceBefore = "0.1cm";
            title.Format.SpaceAfter = "1.5cm";
            title.Format.Font.Name = "Arial";


            //  MDOM.Paragraph sentence = section.AddParagraph("Ce document reporte les indicateurs dosimétriques qui ont été atteints ou non pour le plan de traitement mentionné et approuvé par le Radiothérapeuthe");

            MDOM.Paragraph sentence = section.AddParagraph("Le non respect ponctuel de certaines contraintes de dose sur les organes à risque peut s’avérer nécessaire pour assurer une couverture optimale des volumes cibles, tout en restant dans les limites cliniquement acceptables et en privilégiant le rapport bénéfice-risque pour le patient.");

            sentence.Format.Alignment = MDOM.ParagraphAlignment.Left;
            sentence.Format.Font.Size = 12;
            sentence.Format.Font.Bold = true;
            sentence.Format.Font.Color = MDOM.Colors.Black;
            sentence.Format.SpaceBefore = "0.1cm";
            sentence.Format.SpaceAfter = "1.5cm";
            sentence.Format.Font.Name = "Arial";



            MDOM.Tables.Table table = new MDOM.Tables.Table();
            table.Borders.Width = 1;
            table.Borders.Color = MDOM.Color.FromRgb(200, 200, 200);
            table.AddColumn(MDOM.Unit.FromCentimeter(5));
            table.AddColumn(MDOM.Unit.FromCentimeter(11));

            MDOM.Tables.Row row = table.AddRow();
            MDOM.Tables.Cell cell = row.Cells[0];
            cell.AddParagraph("Patient :");
            cell = row.Cells[1];
            MDOM.Paragraph paragraph = cell.AddParagraph();
            paragraph.AddFormattedText(ctx.Patient.Name);



            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Imprimé le :");
            cell = row.Cells[1];
            paragraph = cell.AddParagraph();
            paragraph.AddFormattedText(DateTime.Now.ToString("g"));

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Approuvé par :");
            cell = row.Cells[1];
            paragraph = cell.AddParagraph();
             docteur = string.Empty;
            string msgdoc = string.Empty;
            if (!pinfo.machine.ToLower().Contains("tom"))
            {
                docteur = pinfo._doctor.Replace(@"admin\", "").Trim().ToUpper();
                msgdoc = "Dr " + docteur;
                msgdoc += " le " + pinfo.approbationDate;
            }
            else
            {
                docteur = pinfo.tomoApproverName;
                msgdoc = "Dr " + docteur;
                msgdoc += " le " + pinfo.tomoApprovalDate;
            }
            paragraph.AddFormattedText(msgdoc);



            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Plan ID (Course ID) :");
            cell = row.Cells[1];
            cell.AddParagraph(ctx.PlanSetup.Id + " (" + ctx.Course.Id + ")");
           

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Machine : ");
            cell = row.Cells[1];
            cell.AddParagraph(pinfo.machine);


            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Technique :");
            cell = row.Cells[1];
            cell.AddParagraph(pinfo.treatmentType);

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Fractionnement principal :");
            cell = row.Cells[1];
            cell.AddParagraph(ctx.PlanSetup.NumberOfFractions.ToString() + " x " + ctx.PlanSetup.DosePerFraction.Dose.ToString("F2") + " Gy");



            section.Add(table);
            #endregion

            #region pdf body



            MDOM.Paragraph paragraph2 = section.AddParagraph("\n\n");
            paragraph2.AddFormattedText("\n", MDOM.TextFormat.Bold);

            List<Ligne> lignes = JsonConvert.DeserializeObject<List<Ligne>>(jsonContent);




            MDOM.Tables.Table table1 = new MDOM.Tables.Table();
            table1.Borders.Width = 1;

            table1.AddColumn(MDOM.Unit.FromCentimeter(7.0));
            table1.AddColumn(MDOM.Unit.FromCentimeter(4.0));
            table1.AddColumn(MDOM.Unit.FromCentimeter(3.0));
            table1.AddColumn(MDOM.Unit.FromCentimeter(2.0));

            row = table1.AddRow();
            //row.Shading.Color = colorMe.PdfTablesTitleBGColor;
            row.Format.Font.Size = 8;
            row.Format.Font.Bold = true;



            row.Cells[0].AddParagraph("Structure");
            row.Cells[1].AddParagraph("Objectif");
            row.Cells[2].AddParagraph("Valeur");
            row.Cells[3].AddParagraph("Atteint ?");

            foreach (Ligne l in lignes)
            {
                bool isok = false;
                bool iswarning = false;
                bool isinfo = false;
                bool iswrong = false;
                if (l.statut.ToLower() == "ok")
                    isok = true;
                else if (l.statut.ToLower() == "var")
                    iswarning = true;
                else if (l.statut.ToLower() == "false")
                    iswrong = true;
                else //if (l.statut.ToLower() == "info")
                    isinfo = true;



                row = table1.AddRow();
                row.Format.Font.Size = 8;
                if (isok)
                {
                    row.Shading.Color = MDOM.Color.FromRgb(0, 220, 0);
                    row.Format.Font.Color = MDOM.Color.FromRgb(0, 0, 0);
                }
                if (iswarning)
                {
                    row.Shading.Color = MDOM.Color.FromRgb(220, 110, 0);
                    row.Format.Font.Color = MDOM.Color.FromRgb(0, 0, 0);

                }
                if (iswrong)
                {
                    row.Shading.Color = MDOM.Color.FromRgb(220, 0, 0);
                    row.Format.Font.Color = MDOM.Color.FromRgb(0, 0, 0);

                }
                if (isinfo)
                {
                    row.Shading.Color = MDOM.Color.FromRgb(220, 220, 220);
                    row.Format.Font.Color = MDOM.Color.FromRgb(0, 0, 0);

                }
                row.Cells[0].AddParagraph(l.structure);
                row.Cells[1].AddParagraph(l.dvhpoint + " " + l.comparateur + " " + l.objectif);
                row.Cells[2].AddParagraph(l.valeur_plan);

                String msg = String.Empty;
                if (isok)
                    msg = "Oui";
                if (iswarning)
                    msg = "Variation";
                if (iswrong)
                    msg = "Non";
                if (isinfo)
                    msg = "Info";
                row.Cells[3].AddParagraph(msg); //.AddFormattedText(msg, TextFormat.Bold); ;


            }


            foreach (Row r in table1.Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    foreach (Paragraph p in c.Elements.OfType<Paragraph>())
                    {
                        p.Format.Alignment = ParagraphAlignment.Center;
                    }
                }
            }

            section.Add(table1);


            #endregion
        }




        public void saveInDirectory()
        {
            #region write pdf
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);

            pdfRenderer.Document = migraDoc;
            pdfRenderer.RenderDocument();
            pdfRenderer.PdfDocument.Save(pdfFilePath);
            #endregion


        }
        public void openThePDF()
        {
            System.Diagnostics.Process.Start(pdfFilePath);
        }
        public void delete()
        {
            if (File.Exists(pdfFilePath))
                File.Delete(pdfFilePath);
        }
        public void saveToAria()
        {


            string _patientId = ctx.Patient.Id;


            InfoUser _infoUser = new InfoUser();
           

            _infoUser.Id = @"admin\"+docteur.ToLower().Trim();// @"admin\vieillevigne"; //_ctx.CurrentUser.Id; // admin\simon_lu
            _infoUser.Language = ctx.CurrentUser.Language; // FRA
            _infoUser.Name = getDoctorName(docteur);// "Vieillevigne Laure";//  _ctx.CurrentUser.Name; // Simon Luc

            //MessageBox.Show("user\n" + _infoUser.Id + "\n" + _infoUser.Language + "\n" + _infoUser.Name);

            string _templateName = ctx.PlanSetup.Id;
            DocumentType _documentType = new DocumentType
            {
                DocumentTypeDescription = "Dosimétrie"  //must be an existing type
            };
            PdfDocument doc = PdfReader.Open(pdfFilePath);
            MemoryStream stream = new MemoryStream();
            doc.Save(stream, false);
            byte[] _binaryContent2 = stream.ToArray();
            DocSettings_2 docSet = DocSettings_2.ReadSettings();
            CustomInsertDocumentsParameter_2.PostDocumentData(_patientId, _infoUser, _binaryContent2, _templateName, _documentType, docSet);


        }


        private string getDoctorName(string doctorId)
        {
            string returnValue = string.Empty;
            string theID = doctorId.Replace(@"admin\", "");
            string userListFilePath = Directory.GetCurrentDirectory() + @"\users\doctors-IUCT.csv";
            if (!File.Exists(userListFilePath))
            {
                MessageBox.Show("Le fichier des utilisateurs est introuvable :\n" + userListFilePath);
                return theID;
            }
            foreach (string line in File.ReadLines(userListFilePath))
            {

                string[] columns = line.Split(';');
                string id = columns[0].ToLower();

                if (id != theID.ToLower())
                    continue;
                else
                {
                    returnValue = columns[2] + " " + columns[1];
                    break;
                }
            }

            if (string.IsNullOrEmpty(returnValue))
            {
                MessageBox.Show("Le nom du docteur n'a pas été trouvé dans le fichier :\n" + userListFilePath + "\nID recherché : " + theID);
                return theID;
            }


            return returnValue.Trim();
        }
    }
    public class InfoUser // because User class from varian is not serializable
    {

        public string Id { get; set; }
        public string Language { get; set; }
        public string Name { get; set; }
    }

}

