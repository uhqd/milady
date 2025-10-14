using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using Newtonsoft.Json;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using VMS.OIS.ARIALocal.WebServices.Document.Contracts;
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

        // Palette douce et imprimable
        private readonly MDOM.Color ColPrimary = MDOM.Color.FromRgb(0x12, 0x5E, 0xA4);   // bleu institutionnel
        private readonly MDOM.Color ColGrey = MDOM.Color.FromRgb(0xF3, 0xF4, 0xF6);      // gris clair
        private readonly MDOM.Color ColBorder = MDOM.Color.FromRgb(0xDD, 0xDD, 0xDD);    // bordure
        private readonly MDOM.Color ColOK = MDOM.Color.FromRgb(0xD9, 0xEF, 0xE3);        // vert pâle
        private readonly MDOM.Color ColWarn = MDOM.Color.FromRgb(0xFF, 0xF3, 0xCD);      // jaune pâle
        private readonly MDOM.Color ColFail = MDOM.Color.FromRgb(0xF8, 0xD7, 0xDA);      // rouge pâle
        private readonly MDOM.Color ColInfo = MDOM.Color.FromRgb(0xEE, 0xEE, 0xEE);      // info neutre

        public createPDF(ScriptContext _ctx, PreliminaryInformation _pinfo, string jsonPath)
        {
            pinfo = _pinfo;
            ctx = _ctx;
            jsonFilePath = jsonPath;

            // --- Chemins & fichiers
            if (!File.Exists(jsonPath))
            {
                MessageBox.Show("Le fichier json n'existe pas:\n" + jsonPath);
                return;
            }
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(jsonPath);
            string pdfDirectory = Path.GetDirectoryName(jsonPath).Replace(@"\json", @"\pdf");
            Directory.CreateDirectory(pdfDirectory); // au cas où
            string dateString = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            pdfFilePath = Path.Combine(pdfDirectory, $"{fileNameWithoutExt}_{dateString}.pdf");

            string jsonContent = File.ReadAllText(jsonPath);

            // Pour extraire IPP, course, plan à partir du nom
            uniqueJsonFileName ujfn = new uniqueJsonFileName();
            ujfn.unmakefilename(Path.GetFileName(jsonPath));
            string ipp = ujfn.IPP;
            string course = ujfn.courseID;
            string plan = ujfn.planID;

            // --- Document & styles
            migraDoc = new MDOM.Document();
            DefineDocumentInfo(ipp, plan);
            DefineStyles(migraDoc);

            MDOM.Section section = migraDoc.AddSection();
            section.PageSetup = new PageSetup
            {
                Orientation = Orientation.Portrait,
                TopMargin = Unit.FromCentimeter(2.0),
                BottomMargin = Unit.FromCentimeter(2.0),
                LeftMargin = Unit.FromCentimeter(2.0),
                RightMargin = Unit.FromCentimeter(2.0)
            };

            AddHeader(section);
            AddTitle(section, "RAPPORT DE LIMITES DE DOSES");

            AddIntro(section);

            AddPatientPlanCard(section, ipp, course, plan);

            // --- Tableau contraintes
            List<Ligne> lignes = JsonConvert.DeserializeObject<List<Ligne>>(jsonContent);
            AddConstraintsTable(section, lignes);

            AddLegend(section);
            AddFooter(section);
        }

       
        public void saveInDirectory()
        {
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = migraDoc;
            pdfRenderer.RenderDocument();
            pdfRenderer.PdfDocument.Save(pdfFilePath);
        }
        public void openThePDF() => System.Diagnostics.Process.Start(pdfFilePath);
        public void delete()
        {
            if (File.Exists(pdfFilePath)) File.Delete(pdfFilePath);
        }
        public void saveToAria()
        {
            string _patientId = ctx.Patient.Id;

            InfoUser _infoUser = new InfoUser
            {
                Id = @"admin\" + docteur.ToLower().Trim(),
                Language = ctx.CurrentUser.Language,
                Name = getDoctorName(docteur)
            };

            string _templateName = ctx.PlanSetup.Id;
            DocumentType _documentType = new DocumentType
            {
                DocumentTypeDescription = "Dosimétrie"
            };

            PdfDocument doc = PdfReader.Open(pdfFilePath);
            byte[] _binaryContent2;
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, false);
                _binaryContent2 = stream.ToArray();
            }
            DocSettings_2 docSet = DocSettings_2.ReadSettings();
            CustomInsertDocumentsParameter_2.PostDocumentData(_patientId, _infoUser, _binaryContent2, _templateName, _documentType, docSet);

        }

        // ---------- Private helpers ----------
        private void DefineDocumentInfo(string ipp, string plan)
        {
            migraDoc.Info.Title = "Rapport de limites de doses";
            migraDoc.Info.Subject = $"Patient {ipp} – Plan {plan}";
            migraDoc.Info.Author = "Service de Radiothérapie";
        }

        private void DefineStyles(MDOM.Document doc)
        {
            // Police par défaut (sobre et hospitalière)
            MDOM.Style normal = doc.Styles["Normal"];
            normal.Font.Name = "Arial";
            normal.Font.Size = 9;

            // Titre principal
            MDOM.Style h1 = doc.Styles["Heading1"];
            h1.Font.Name = "Arial";
            h1.Font.Size = 16;
            h1.Font.Bold = true;
            h1.Font.Color = ColPrimary;
            h1.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            h1.ParagraphFormat.SpaceBefore = Unit.FromPoint(6);
            h1.ParagraphFormat.SpaceAfter = Unit.FromPoint(10);

            // Sous-titres
            MDOM.Style h2 = doc.Styles["Heading2"];
            h2.Font.Name = "Arial";
            h2.Font.Size = 11.5;
            h2.Font.Bold = true;
            h2.Font.Color = MDOM.Colors.Black;
            h2.ParagraphFormat.SpaceBefore = Unit.FromPoint(12);
            h2.ParagraphFormat.SpaceAfter = Unit.FromPoint(6);

            // Style des “labels” (cartouche)
            var label = doc.Styles.AddStyle("Label", "Normal");
            label.Font.Bold = true;

            // Style tableau générique
            var t = doc.Styles.AddStyle("TableBase", "Normal");
            t.ParagraphFormat.SpaceBefore = Unit.FromPoint(2);
            t.ParagraphFormat.SpaceAfter = Unit.FromPoint(2);
        }

        private void AddHeader(MDOM.Section section)
        {
            string logoPath = Path.Combine(Directory.GetCurrentDirectory(), "img", "logo.jpg");

            // En-tête principal
            var header = section.Headers.Primary;

            // 1) Logo (facultatif si le fichier est absent)
            if (File.Exists(logoPath))
            {
                var img = header.AddImage(logoPath);
                img.Height = Unit.FromCentimeter(1.0);
                img.LockAspectRatio = true;
                header.AddParagraph().AddLineBreak(); // ajoute une petite marge sous le logo
            }

            

            // 2) Liseré de séparation : un paragraphe avec bordure basse
            var sep = header.AddParagraph();
           // sep.Format.SpaceBefore = MDOM.Unit.FromPoint(3);
           // sep.Format.SpaceAfter = MDOM.Unit.FromPoint(0);
            sep.Format.SpaceBefore = Unit.FromCentimeter(1.3); // espace sous le logo
            sep.Format.SpaceAfter = Unit.FromCentimeter(0.8);

            sep.Format.Borders.Bottom.Color = ColBorder;  
            sep.Format.Borders.Bottom.Width = 0.75;
            // petit espace pour forcer le rendu de la bordure
            sep.AddText(" ");
            section.PageSetup.HeaderDistance = Unit.FromCentimeter(2.5);

        }

        private void AddFooter(MDOM.Section section)
        {
            var footer = section.Footers.Primary;
            var p = footer.AddParagraph();
            p.AddFormattedText("Document interne – Dossier patient de radiothérapie  - " + ctx.Patient.Name, TextFormat.Italic);
            p.Format.Font.Size = 8;
            p.Format.Alignment = ParagraphAlignment.Left;

            var p2 = footer.AddParagraph();
            p2.AddText("Page ");
            p2.AddPageField();
            p2.AddText(" / ");
            p2.AddNumPagesField();
            p2.Format.Font.Size = 8;
            p2.Format.Alignment = ParagraphAlignment.Right;
        }

        private void AddTitle(MDOM.Section section, string title)
        {
            var t = section.AddParagraph(title, "Heading1");
            t.Format.SpaceBefore = Unit.FromCentimeter(2.0);  // espace avant le titre
            t.Format.SpaceAfter = Unit.FromCentimeter(1.8);  // léger espace après
        }

        private void AddIntro(MDOM.Section section)
        {
            var s = section.AddParagraph();
            s.Format.SpaceBefore = Unit.FromPoint(6);
            s.Format.SpaceAfter = Unit.FromPoint(10);
            s.Format.Alignment = ParagraphAlignment.Left;
            s.Format.Borders.Bottom.Color = ColBorder;
            s.Format.Borders.Bottom.Width = 0.5;

            s.AddFormattedText(
                "Le non-respect ponctuel de certaines contraintes de dose sur les organes à risque peut s’avérer nécessaire " +
                "pour assurer une couverture optimale des volumes cibles, dans des limites cliniquement acceptables, " +
                "en privilégiant le rapport bénéfice–risque pour le patient.",
                TextFormat.NotBold);
        }

        private void AddPatientPlanCard(MDOM.Section section, string ipp, string course, string plan)
        {
            section.AddParagraph("Informations patient et plan", "Heading2");

            var table = new MDOM.Tables.Table { Style = "TableBase" };
            table.Borders.Color = ColBorder;
            table.Borders.Width = 0.75;
            table.TopPadding = 3;
            table.BottomPadding = 3;

            table.AddColumn(Unit.FromCentimeter(5.0));
            table.AddColumn(Unit.FromCentimeter(11.0));

            AddKeyValueRow(table, "Patient :", Safe(ctx?.Patient?.Name));
            AddKeyValueRow(table, "IPP :", ipp);
            AddKeyValueRow(table, "Imprimé le :", DateTime.Now.ToString("g", new CultureInfo("fr-FR")));

            // Médecin approbateur
            string msgdoc;
            if (!pinfo.machine?.ToLower()?.Contains("tom") ?? false)
            {
                docteur = (pinfo._doctor ?? "").Replace(@"admin\", "").Trim().ToUpperInvariant();
                msgdoc = $"Dr {docteur} le {pinfo.approbationDate}";
            }
            else
            {
                docteur = pinfo.tomoApproverName;
                msgdoc = $"Dr {docteur} le {pinfo.tomoApprovalDate}";
            }
            AddKeyValueRow(table, "Approuvé par :", msgdoc);

            // Plan / Course
            AddKeyValueRow(table, "Plan ID (Course ID) :", $"{Safe(ctx?.PlanSetup?.Id)} ({Safe(ctx?.Course?.Id)})");
            AddKeyValueRow(table, "Machine :", Safe(pinfo.machine));
            AddKeyValueRow(table, "Technique :", Safe(pinfo.treatmentType));

            // Fractionnement
            string fracTxt = BuildFractionnement();
            AddKeyValueRow(table, "Fractionnement principal :", fracTxt);

            section.Add(table);
        }

        private void AddConstraintsTable(MDOM.Section section, List<Ligne> lignes)
        {
            section.AddParagraph("Contraintes dosimétriques", "Heading2");

            var table = new MDOM.Tables.Table { Style = "TableBase" };
            table.Borders.Width = 0.75;
            table.Borders.Color = ColBorder;
            table.Rows.LeftIndent = 0;

            table.AddColumn(Unit.FromCentimeter(7.0)); // Structure
            table.AddColumn(Unit.FromCentimeter(4.5)); // Objectif
            table.AddColumn(Unit.FromCentimeter(3.0)); // Valeur
            table.AddColumn(Unit.FromCentimeter(1.5)); // Atteint ?

            var header = table.AddRow();
            header.HeadingFormat = true;
            header.Shading.Color = ColGrey;
            header.Format.Font.Bold = true;
            header.Format.Alignment = ParagraphAlignment.Center;
            header.Format.Font.Size = 9;
            header.Cells[0].AddParagraph("Structure");
            header.Cells[1].AddParagraph("Objectif");
            header.Cells[2].AddParagraph("Valeur");
            header.Cells[3].AddParagraph("Atteint ?");

            if (lignes != null)
            {
                foreach (var l in lignes)
                {
                    var (isok, iswarning, iswrong, isinfo) = MapStatus(l.statut);
                    var row = table.AddRow();
                    row.Format.Font.Size = 9;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;

                    if (isok) row.Shading.Color = ColOK;
                    else if (iswarning) row.Shading.Color = ColWarn;
                    else if (iswrong) row.Shading.Color = ColFail;
                    else row.Shading.Color = ColInfo;

                    row.Cells[0].AddParagraph(Safe(l.structure));
                    row.Cells[0].Format.Alignment = ParagraphAlignment.Left;

                    row.Cells[1].AddParagraph($"{Safe(l.dvhpoint)} {Safe(l.comparateur)} {Safe(l.objectif)}");
                    row.Cells[1].Format.Alignment = ParagraphAlignment.Center;

                    row.Cells[2].AddParagraph(Safe(l.valeur_plan));
                    row.Cells[2].Format.Alignment = ParagraphAlignment.Center;

                    row.Cells[3].AddParagraph(StatusLabel(isok, iswarning, iswrong, isinfo));
                    row.Cells[3].Format.Alignment = ParagraphAlignment.Center;
                }
            }

            // Harmoniser l’alignement du header
            foreach (Cell c in table.Rows[0].Cells)
                foreach (Paragraph p in c.Elements.OfType<Paragraph>())
                    p.Format.Alignment = ParagraphAlignment.Center;

            section.Add(table);
        }

        private void AddLegend(MDOM.Section section)
        {
            var p = section.AddParagraph();
            p.Format.SpaceBefore = Unit.FromPoint(6);

            var t = new MDOM.Tables.Table();
            t.Borders.Color = ColBorder;
            t.Borders.Width = 0.5;
            t.AddColumn(Unit.FromCentimeter(3.0));
            t.AddColumn(Unit.FromCentimeter(13.0));

            AddLegendRow(t, ColOK, "Oui", "Objectif atteint.");
            AddLegendRow(t, ColWarn, "Variation", "Écart mineur / acceptable selon contexte clinique.");
            AddLegendRow(t, ColFail, "Non", "Objectif non atteint");
            AddLegendRow(t, ColInfo, "Info", "Information sans objectif (ou non applicable).");

            section.AddParagraph("Légende", "Heading2");
            section.Add(t);
        }

        private void AddLegendRow(MDOM.Tables.Table t, MDOM.Color shade, string label, string text)
        {
            var r = t.AddRow();
            r.Shading.Color = shade;
            r.Cells[0].AddParagraph(label).Format.Font.Bold = true;
            r.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            r.Cells[1].AddParagraph(text);
        }

        private void AddKeyValueRow(MDOM.Tables.Table table, string key, string value)
        {
            var row = table.AddRow();
            row.TopPadding = 2;
            row.BottomPadding = 2;
            row.Cells[0].AddParagraph(key).Style = "Label";
            row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].AddParagraph(value);
            row.Cells[1].Format.Alignment = ParagraphAlignment.Left;

            // ligne séparatrice légère
            row.Borders.Bottom.Color = ColBorder;
            row.Borders.Bottom.Width = 0.5;
        }

        private string BuildFractionnement()
        {
            try
            {
                var n = ctx?.PlanSetup?.NumberOfFractions;
                var d = ctx?.PlanSetup?.DosePerFraction.Dose;
                if (n.HasValue && d.HasValue)
                    return $"{n} × {d.Value:F2} Gy";
            }
            catch { /* ignore */ }
            return "-";
        }

        private static string Safe(string s) => string.IsNullOrWhiteSpace(s) ? "-" : s;

        private (bool ok, bool warn, bool fail, bool info) MapStatus(string statut)
        {
            if (string.IsNullOrWhiteSpace(statut))
                return (false, false, false, true);

            var st = statut.Trim().ToLowerInvariant();

            switch (st)
            {
                case "ok":
                    return (true, false, false, false);
                case "var":
                    return (false, true, false, false);
                case "false":
                case "non":
                    return (false, false, true, false);
                default:
                    return (false, false, false, true);
            }
        }


        private string StatusLabel(bool ok, bool warn, bool fail, bool info)
        {
            if (ok) return "Oui";
            if (warn) return "Variation";
            if (fail) return "Non";
            return "Info";
        }

        private string getDoctorName(string doctorId)
        {
            string theID = (doctorId ?? "").Replace(@"admin\", "");
            string userListFilePath = Path.Combine(Directory.GetCurrentDirectory(), "users", "doctors-IUCT.csv");
            if (!File.Exists(userListFilePath))
            {
                MessageBox.Show("Le fichier des utilisateurs est introuvable :\n" + userListFilePath);
                return theID;
            }
            foreach (string line in File.ReadLines(userListFilePath))
            {
                string[] columns = line.Split(';');
                if (columns.Length < 3) continue;
                string id = columns[0].ToLowerInvariant();
                if (id != theID.ToLowerInvariant()) continue;
                return (columns[2] + " " + columns[1]).Trim();
            }

            MessageBox.Show("Le nom du docteur n'a pas été trouvé dans :\n" + userListFilePath + "\nID recherché : " + theID);
            return theID;
        }
    }

    public class InfoUser
    {
        public string Id { get; set; }
        public string Language { get; set; }
        public string Name { get; set; }
    }
}
