using Newtonsoft.Json;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;

using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using VMS.OIS.ARIAExternal.WebServices.Documents.Contracts;
using VMS.TPS;
using VMS.TPS.Common.Model;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Windows.Forms;

namespace catchDose
{
    internal class docFinder
    {
        public string ApproverName { get; set; }
        public string ApprovalDate { get; set; }
        private List<DateTime> dosimetrie = new List<DateTime>();
        private bool myPlanReportIsFound;
        private docExtractor _pprc;
        public bool failedToFind { get; set; }
        public docFinder(ScriptContext ctx)
        {

            failedToFind = false;
            String response = connectToAriaDocuments(ctx);

            if (response != null)
            {


                parseTheAriaDocuments(response, ctx); // check if documents exists and get info from tomo report if needed

            }
            else
            {
                System.Windows.MessageBox.Show("Aucun document trouvé dans Aria Documents");
                failedToFind = true;
            }


        }

        public String connectToAriaDocuments(ScriptContext ctx) // connect to ARIA, return request response
        {
            bool DocumentAriaIsConnected = true;
            ServicePointManager.ServerCertificateValidationCallback += (o, c, ch, er) => true;

            #region get apikey, host, port
            DocSettings docSet = DocSettings.ReadSettings();
            string apiKeyDoc = docSet.DocKey;
            string hostName = docSet.HostName;
            string port = docSet.Port;
            #endregion

            string response = "";
            string request = "{\"__type\":\"GetDocumentsRequest:http://services.varian.com/Patient/Documents\",\"Attributes\":[],\"PatientId\":{ \"ID1\":\"" + ctx.Patient.Id + "\"}}";
            try
            {
                response = CustomInsertDocumentsParameter.SendData(request, true, apiKeyDoc, docSet.HostName, docSet.Port);
            }
            catch
            {
                System.Windows.MessageBox.Show("La connexion à Aria Documents a échoué. Les documents ne peuvent pas être récupérés");

                DocumentAriaIsConnected = false;
            }

            if (DocumentAriaIsConnected)
            {
                //MessageBox.Show("Aria Connected");
                return response;
            }
            else
            {
                //MessageBox.Show("Aria not Connected");
                return null;
            }
        }
        public void parseTheAriaDocuments(String response, ScriptContext ctx) // using the request response, parse the documents
        {

            #region declaration of variables and deserialize response

            DocSettings docSet = DocSettings.ReadSettings();
            ServicePointManager.ServerCertificateValidationCallback += (o, c, ch, er) => true;
            string apiKeyDoc = docSet.DocKey;
            string hostName = docSet.HostName;
            string port = docSet.Port;

            //string doc1 = "Dosimétrie";

            var VisitNoteList = new List<string>();
            int visitnoteloc = response.IndexOf("PtVisitNoteId");
            while (visitnoteloc > 0)
            {
                VisitNoteList.Add(response.Substring(visitnoteloc + 15, 2).Replace(",", ""));
                visitnoteloc = response.IndexOf("PtVisitNoteId", visitnoteloc + 1);
            }
            var response_Doc = JsonConvert.DeserializeObject<DocumentsResponse>(response); // get the list of documents
            var DocTypeList = new List<string>();
            var DateServiceList = new List<DateTime>();
            List<int> DocIndexList = new List<int>();
            var PatNameList = new List<string>();
            int loopnum = 0;


            string thePtId = "";
            string thePtVisitId = "";
            string theVisitNoteId = "";
            string request_docdetails = "";
            string response_docdetails = "";
            string thisDocType = "";
            string thisDocName = "x";
            int typeloc = 0;
            int enteredloc = 0;
            int templateloc = 0;
            System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);

            #endregion

            #region loop on documents to select document of interest

            int indexer = 0;

            foreach (var document in response_Doc.Documents) // parse documents
            {


                bool sendItToTrash = false;


                #region get the general infos of this document , dismiss if date of service <0
                thePtId = document.PtId;
                thePtVisitId = document.PtVisitId.ToString();
                theVisitNoteId = VisitNoteList[loopnum];
                request_docdetails = "{\"__type\":\"GetDocumentRequest:http://services.varian.com/Patient/Documents\",\"Attributes\":[],\"PatientId\":{ \"PtId\":\"" + thePtId + "\"},\"PatientVisitId\":" + thePtVisitId + ",\"VisitNoteId\":" + theVisitNoteId + "}";
                response_docdetails = CustomInsertDocumentsParameter.SendData(request_docdetails, true, apiKeyDoc, docSet.HostName, docSet.Port);
                typeloc = response_docdetails.IndexOf("DocumentType");
                enteredloc = response_docdetails.IndexOf("EnteredBy");
                templateloc = response_docdetails.IndexOf("TemplateName");



                if (typeloc > 0)
                {
                    thisDocType = response_docdetails.Substring(typeloc + 15, enteredloc - typeloc - 18);

                }
                if (templateloc > 0)
                {
                    int mlength = response_docdetails.Length - templateloc - 15 - 3;
                    thisDocName = response_docdetails.Substring(templateloc + 15, mlength);// response_docdetails.Length - 3);

                }

                int nameloc = response_docdetails.IndexOf("PatientLastName");
                int dobloc = response_docdetails.IndexOf("PreviewText");
                if (nameloc > 0)
                    PatNameList.Add(response_docdetails.Substring(nameloc + 18, dobloc - nameloc - 21));
                int dateservloc = response_docdetails.IndexOf("DateOfService");
                int datesignloc = response_docdetails.IndexOf("DateSigned");

                if (dateservloc <= 0)
                {
                    sendItToTrash = true;

                }
                #endregion


                #region  dismiss if document is marked as error
                if (!sendItToTrash)
                {
                    int IsMarkedAsErrorIndex = response_docdetails.IndexOf("IsMarkedAsError"); // TRUE = ERROR   FALSE = OK :-)
                    string isError = response_docdetails.Substring(IsMarkedAsErrorIndex + 17, 4);

                    if (isError.ToUpper().Contains("TRU"))
                    {
                        sendItToTrash = true;

                    }

                }
                #endregion

                #region getdate                
                if (!sendItToTrash)
                {

                    dtDateTime = dtDateTime.AddSeconds(Convert.ToDouble(response_docdetails.Substring(dateservloc + 23, datesignloc - dateservloc - 34)) / 1000).ToLocalTime();




                }
                #endregion

                #region dismiss if document has a useless type 
                if (!sendItToTrash)
                {


                    //   if ((thisDocType != doc1)) // must be Dosimétrie
                    //{
                    //    sendItToTrash = true;
                    //  }
                }
                #endregion




                indexer++;
                #region If not dismissed, add the document to the list
                if (!sendItToTrash)
                {
                  //  if (thisDocType.Contains("Dosimétrie"))
                   // {
                        docExtractor docExtractor = new docExtractor(response_docdetails, thisDocType, thisDocName,ctx, indexer);


                        DateServiceList.Add(dtDateTime);
                        DocTypeList.Add(thisDocType);
                        DocIndexList.Add(loopnum);
                    //}
                }
                #endregion


                loopnum++;
            }
            #endregion







        }

    }
    public class docExtractor
    {
        private tomoReportData trda;
        private eclipseReportData erd;

        private string outpath = "";

        public bool isAPlanReport { get; set; }
        public bool isATomoReport { get; set; }
        public bool isAEclipseReport { get; set; }
        public bool isTheCorrectTomoReport { get; set; }
        public bool isTheCorrectEclipseReport { get; set; }
        public bool isTheCorrectEclipseReportWrongDate { get; set; }

        public docExtractor(string response_details, string doctype, string docname,ScriptContext _ctx, int indexer)
        {
            bool typeIsKnown = true;
            string extension = ".pdf";
            if (doctype.Contains("Dosimétrie"))
                extension = ".pdf";
            else if (doctype.Contains("Fiche de positionnement"))
                extension = ".docx";
            else
                typeIsKnown= false;

            if (typeIsKnown)
            {
                #region Convert response details (string) to a temp PDF file
                String saveFilePathDir = @"\\srv015\sf_com\simon_lu\milady";
                saveFilePathDir += @"\" + _ctx.Patient.Id + "_" + _ctx.Patient.LastName;

//                String saveFilePathDir = Directory.GetCurrentDirectory() + @"\" + _ctx.Patient.Id + "_" + _ctx.Patient.LastName;
                if (!Directory.Exists(saveFilePathDir))
                    Directory.CreateDirectory(saveFilePathDir);

                string cleanType = doctype.Trim().Replace(" ", "");
                string cleanName = docname.Trim().Replace(" ", "");
                String saveFilePathTemp = saveFilePathDir + @"\" + indexer +"_"+ cleanType+"_"+ cleanName +extension;
                String saveFilePathTempresponse = saveFilePathDir + @"\" + indexer + "__temp__.txt" ;

                int startBinary = response_details.IndexOf("\"BinaryContent\"") + 17;
                int endBinary = response_details.IndexOf("\"Certifier\"") - 2;
                string binaryContent2 = response_details.Substring(startBinary, endBinary - startBinary);
                binaryContent2 = binaryContent2.Replace("\\", "");  // the \  makes the string a non valid base64 string                       
                File.WriteAllBytes(saveFilePathTemp, Convert.FromBase64String(binaryContent2));

              //  File.WriteAllText(saveFilePathTempresponse, response_details);
                #endregion

            }


            // if (File.Exists(saveFilePathTemp))
            //   File.Delete(saveFilePathTemp);




        }


        private void readThePDF(string pathToPdf, ScriptContext ctx)
        {

            trda = new tomoReportData();
            erd = new eclipseReportData();


            #region convert pdf 2 text file and detect if Eclipse or tomo
            string dateString = DateTime.Now.ToString().Replace(" ", "").Replace(":", "_").Replace("/", "-").Replace(@"\", "-");
            outpath = Directory.GetCurrentDirectory() + @"\tomoReportData_" + ctx.CurrentUser.Name.Replace(" ", "") + dateString + ".txt";


            trda.planReportDataAsATextFilePath = outpath;

            String pageContent = readWithPdfPig(pathToPdf);

            File.WriteAllText(outpath, pageContent, Encoding.UTF8);


            if (pageContent.Contains("Accuray"))  /// remove old tomo plan
            {
                if (!pageContent.Contains("Accuray Precision 1.1.1.0"))
                    if (!pageContent.Contains("Accuray Precision 2.0.0.1"))
                        isATomoReport = true;
            }
            else if (pageContent.Contains("appportdeplanificationderadiothérapie")) // FX printer script
            {
                isAEclipseReport = true;
            }

            //File.WriteAllText(outpath, pageContent);
            // pdfDoc.Close();
            // pdfReader.Close();
            #endregion

            if (isATomoReport)
            {

                #region read text file in a list of strings
                StreamReader file = new StreamReader(trda.planReportDataAsATextFilePath);
                String line = null;
                List<string> lines = new List<string>();

                while ((line = file.ReadLine()) != null)
                {
                    lines.Add(line);

                }

                #endregion


                #region Get the infos (see ex. at the end of file)




                string[] separatingStrings = { "rev" };
                string[] separatingStrings2 = { ":" };
                string[] separatingStrings3 = { ", " };
                string[] separatingStrings4 = { "of " };
                bool planNameFound = false;  // because Plan name is several times in the file
                bool patientNameFound = false;  // because patient name is several times in the file





                //for (int i = 0; i < lines.Count; i++)

                foreach (string lne in lines)
                {
                    /*
                    if ((lne.Contains("Plan Name:")) && (!planNameFound))
                    {
                        trd.planName = lines[i + 2];
                        planNameFound = true;
                    }
                    */
                    /*
                    if (lne.Contains("Plan Type:"))
                        trd.planType = lines[i + 2];
                    */
                    /*
                    if (lne.Contains("Treatment Machine"))
                    {
                        trd.machineNumber = lines[i + 1];
                        string[] sub1 = lne.Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries);
                        string[] sub2 = sub1[1].Split('/');
                        trd.machineRevision = sub2[0];
                    }
                    */
                    /* if (lne.Contains("Prescription:")) //Prescription: Median of PTV sein, 50.00 Gy
                     {
                         string[] sub1 = lne.Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries); // Prescription ... Median of PTV sein, 50.00 Gy
                         string[] sub2 = sub1[1].Split(separatingStrings3, System.StringSplitOptions.RemoveEmptyEntries); // Median of PTV sein ... 50.00 Gy
                         string[] sub3 = sub2[1].Split(' ');                      // 
                         trd.prescriptionMode = sub2[0];
                         trd.prescriptionTotalDose = Convert.ToDouble(sub3[0]);
                         trd.prescriptionStructure = sub2[0].Split(separatingStrings4, System.StringSplitOptions.RemoveEmptyEntries)[1];
                         trd.prescriptionMode = sub2[0].Split(separatingStrings4, System.StringSplitOptions.RemoveEmptyEntries)[0];
                     }
                     */
                    /*
                     if (lne.Contains("Prescribed Dose per Fraction"))
                         trd.prescriptionDosePerFraction = Convert.ToDouble(lne.Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries)[1]);

                     if (lne.Contains("Planned Fractions"))
                         trd.prescriptionNumberOfFraction = Convert.ToInt32(lne.Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries)[1]);
                     */


                    if (lne.Contains("Maximum Dose"))
                        try
                        {
                            string temp = lne.Replace(" ", "").Trim();
                            trda.maxDose = Convert.ToDouble(temp.Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries)[2]);
                            // System.Windows.MessageBox.Show("max dose is " + trda.maxDose);
                        }
                        catch
                        {
                            System.Windows.MessageBox.Show("Erreur lors de la lecture de la dose maximale dans le rapport PDF. Veuillez vérifier le format du rapport : " + lne);
                        }

                    if (lne.Contains("Referring Physician:"))
                        try
                        {
                            string[] sub2 = lne.Split(',');
                            string[] sub3 = sub2[0].Split(':');
                            trda.approvalDate = sub3[sub3.Length - 1].Trim();


                            //trda.approvalDate = lne.Replace("Plan Saved Date: ", "").Split(',')[0];
                            // System.Windows.MessageBox.Show("approve date is " + trda.approvalDate);
                        }
                        catch
                        {
                            System.Windows.MessageBox.Show("Erreur lors de la lecture de la date dans le rapport PDF. Veuillez vérifier le format du rapport : " + lne);
                        }
                    /*
                    if (lne.Contains("Plan Status:"))
                    {

                        try
                        {
                            trd.approvalStatus = lne.Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries)[1];
                        }
                        catch
                        {
                            trd.approvalStatus = "non_trouvé";
                        }
                    }*/

                    /*
                    if (lne.Contains("MU per Fraction:"))
                    {
                        string[] sub2 = lne.Split(':');
                        string[] sub3 = sub2[1].Split('/');
                        trd.MUplanned = Convert.ToDouble(sub3[0]);
                        trd.MUplannedPerFraction = Convert.ToDouble(sub3[1]);

                    }

                    if (lne.Contains("Field Width")) //Field Width (cm): 5.0, Dynamic
                    {

                        string[] sub2 = lne.Split(':');
                        string[] sub3 = sub2[1].Split(',');
                        trd.fieldWidth = Convert.ToDouble(sub3[0]);

                        if (sub3[1].Contains("Dynamic"))
                            trd.isDynamic = true;
                        else
                            trd.isDynamic = false;
                    }

                    if (lne.Contains("Pitch:"))
                    {

                        string[] sub2 = lne.Split(':');
                        trd.pitch = Convert.ToDouble(sub2[1]);
                    }
                    if (lne.Contains("Modulation Factor"))
                    {

                        string[] sub2 = lne.Split(':');
                        string[] sub3 = sub2[1].Split('/');
                        trd.modulationFactor = Convert.ToDouble(sub3[0]);

                    }
                    if (lne.Contains("Gantry Period (sec)"))
                    {

                        try
                        {
                            trd.gantryPeriod = Convert.ToDouble(lines[i + 2]);
                            trd.gantryNumberOfRotation = Convert.ToDouble(lines[i + 3]);
                        }
                        catch
                        {
                            trd.gantryPeriod = 99.999;
                            trd.gantryNumberOfRotation = 99.999;

                        }
                    }
                    if (lne.Contains("Couch Travel (mm)"))
                    {
                        try
                        {
                            trd.couchTravel = Convert.ToDouble(lines[i + 2]);
                            trd.couchSpeed = Convert.ToDouble(lines[i + 3]);
                        }
                        catch
                        {
                            trd.couchTravel = 99.99;
                            trd.couchSpeed = 99.99;
                        }
                    }

     
                    if (lne.Contains("Red Lasers Offset (IECf, mm)"))
                    {
                        string[] sub2 = lne.Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries); // split with ": "

                        string[] sub3 = sub2[2].Split(' ');
                        trd.redLaserXoffset = Convert.ToDouble(sub3[0]);

                        string[] sub4 = sub2[3].Split(' ');
                        trd.redLaserYoffset = Convert.ToDouble(sub4[0]);

                        string[] sub5 = sub2[4].Split(' ');
                        trd.redLaserZoffset = Convert.ToDouble(sub5[0]);
                    }


                    if (lne.Contains("Exit Only"))
                    {
                        trd.blockedOAR.Add(getStructureNameInTheLine(lne));


                    }

                    if (lne.Contains("Beam On Time"))
                    {
                        string[] sub2 = lne.Split(':');
                        trd.beamOnTime = Convert.ToDouble(sub2[1]);
                    }
                    if (lne.Contains("Delivery Type"))
                    {

                        trd.deliveryMode = lne;
                    }
                    if (lne.Contains("Dose Calculation Algorithm:"))//Convolution-Superposition Spacing (IECp, mm) X: 1.27 Y: 2.50 Z: 1.27
                    {
                        string[] sub2 = lines[i + 2].Split('(');
                        string[] sub2b = sub2[0].Split(' ');
                        trd.algorithm = sub2b[0];

                        string[] sub3 = lines[i + 1].Split('/');
                        string[] sub4 = sub3[1].Split(':');
                        trd.resolutionCalculation = sub4[1].Trim(); // remove space char

                        string[] sub5 = lines[i + 2].Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries); // split with ": "
                        string[] sub6 = sub5[1].Split(' ');
                        trd.refPointX = Convert.ToDouble(sub6[0]);
                        string[] sub7 = sub5[2].Split(' ');
                        trd.refPointY = Convert.ToDouble(sub7[0]);
                        string[] sub8 = sub5[3].Split(' ');
                        trd.refPointZ = Convert.ToDouble(sub8[0]);
                    }
                    */
                    /* if (lne.Contains("Reference Dose (Gy)"))
                     {
                         try
                         {
                             string[] sub2 = lne.Split(':');
                         trd.refDose = Convert.ToDouble(sub2[1]);
                         }
                         catch
                         {
                             MessageBox.Show("Erreur lors de la lecture de la dose ref dans le rapport PDF. Veuillez vérifier le format du rapport : " + lne);
                         }
                     }*/
                    /*
                    if (lne.Contains("Planning Method:")) //Planning Method: Classic
                    {
                        string[] sub2 = lne.Split(':');
                        trd.planningMethod = sub2[1];

                    }
                    if (!patientNameFound)
                        if ((lne.Contains("Patient Name:")) && (lne.Contains("Plan Name:"))) //Patient Name: TOTO, TITI; Medical ID: 123456789; Plan Name: SeinG+gg; Version: Accuray Precision 
                        {
                            patientNameFound = true;

                            string[] sub2 = lne.Split(';');//Patient Name: TOTO, TITI    +     Medical ID: 123456789   +   ...


                            string[] sub3 = sub2[0].Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries); // split with ": "
                            string[] sub4 = sub3[1].Split(separatingStrings3, System.StringSplitOptions.RemoveEmptyEntries); // split with ", "
                            trd.patientName = sub4[0];
                            trd.patientFirstName = sub4[1];

                            string[] sub5 = sub2[1].Split(separatingStrings2, System.StringSplitOptions.RemoveEmptyEntries); // split with ": "
                            trd.patientID = sub5[1];

                        }

                    if (lne.Contains("Density Model:"))
                    {
                        string[] sub2 = lne.Split(':');
                        trd.HUcurve = sub2[1];
                    }
               */
                    //if (lne.Contains("User ID:"))
                    if (lne.Contains(" Radiation Oncologist:"))
                    {
                        try
                        {
                            string[] sub2 = lne.Split(',');
                            string[] sub3 = sub2[0].Split(' ');
                            trda.approverID = sub3[sub3.Length - 1];
                            // System.Windows.MessageBox.Show("approver  is " + trda.approverID);

                        }
                        catch
                        {
                            System.Windows.MessageBox.Show("Erreur lors de la lecture du RT dans le rapport PDF. Veuillez vérifier le format du rapport : " + lne);
                        }
                    }

                    /*
                    if ((lne.Contains("Patient Position:")) && (!lne.Contains("Delivery"))) //Planning Method: Classic
                    {
                        trd.patientPosition = lines[i + 2];
                    }


                    trd.originX = 99999;
                    if (lne.Contains("Origin")) //Origin(IECp, mm) -325.000 -122.500 325.000
                    {

                        string[] sub2 = lne.Split(' ');
                        trd.originX = Convert.ToDouble(sub2[3]);
                        trd.originY = Convert.ToDouble(sub2[4]);
                        trd.originZ = Convert.ToDouble(sub2[5]);
                    }

                    if (lne.Contains("Scan Date"))
                    {

                        string[] sub2 = lines[i + 3].Split(',');
                        trd.CTDate = sub2[0];
                    }
                    */


                }
                //checkTRD(trd);


                #endregion

            }
            /*
            else if (isAEclipseReport)
            {

                #region read text file in a list of strings
                System.IO.StreamReader file = new System.IO.StreamReader(outpath);
                String line = null;
                List<string> lines = new List<string>();

                while ((line = file.ReadLine()) != null)
                {
                    lines.Add(line);

                }

                #endregion

                for (int i = 0; i < lines.Count; i++)
                {
                    #region get the approbation date in the plan report
                    try
                    {
                        if ((lne.Contains("Appro.plan:"))) // ex. Appro.plan: KellerAudrey(vendredi5juillet202416:57:00)
                        {
                            String brutedate = lne;
                            string pattern = @"\(([^)]+)\)";
                            Match match = Regex.Match(brutedate, pattern);

                            if (!match.Success) // sometimes date is on two lines
                            {
                                brutedate = lne + lines[i + 1];
                                match = Regex.Match(brutedate, pattern);
                            }

                            if (match.Success)
                            {
                                string dateTimeString = match.Groups[1].Value;
                                // string format: "vendredi5juillet202416:57:00"

                                pattern = @"(\D+)(\d+)(\D+)(\d{4})(\d{2}:\d{2}:\d{2})";
                                match = Regex.Match(dateTimeString, pattern);


                                if (match.Success)
                                {
                                    string dayOfWeek = match.Groups[1].Value;
                                    string day = match.Groups[2].Value;

                                    string month = match.Groups[3].Value;
                                    string year = match.Groups[4].Value;
                                    string time = match.Groups[5].Value;
                                    string formattedDateTime = $"{day} {month} {year} {time}";


                                    var cultureInfo = new CultureInfo("fr-FR");
                                    var dateTimeFormat = cultureInfo.DateTimeFormat;
                                    DateTime.TryParseExact(formattedDateTime, "d MMMM yyyy HH:mm:ss", dateTimeFormat, DateTimeStyles.None, out erd.approDate);

                                }
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Erreur lors de la lecture de la date d'approbation du plan dans le rapport PDF. Veuillez vérifier le format du rapport : " + lne);
                    }
                    #endregion



                    #region get patientID in the plan report
                    try
                    {
                        if (lne.Contains("Nom") && lne.Contains("Date") && lne.Contains("Sexe")) // first line in fx report
                        {

                            string[] sub2 = lne.Split('(');
                            string[] sub3 = sub2[1].Split(')');
                            erd.patientID = sub3[0];

                        }
                    }
                    catch
                    {
                        MessageBox.Show("Erreur lors de la lecture du patientID. Veuillez vérifier le format du rapport : " + lne);
                    }
                    #endregion


                    #region get courseID in the plan report
                    try
                    {
                        if (!string.IsNullOrEmpty(lne) && lne.Length >= 2 && lne.StartsWith("ID"))
                            if (!lne.Contains("plan"))
                                if (lne.Contains("ID") && lne.Contains("dossier") && lne.Contains("traitement"))
                                {
                                    string[] sub2 = lne.Split(' ');

                                    erd.courseID = sub2[1];
                                }
                    }
                    catch
                    {
                        MessageBox.Show("Erreur lors de la lecture du courseID. Veuillez vérifier le format du rapport : " + lne);
                    }

                    #endregion


                    #region get planID in the plan report
                    try
                    {
                        if (!string.IsNullOrEmpty(lne) && lne.Length >= 2 && lne.StartsWith("ID"))
                            if (lne.Contains("ID") && lne.Contains("plan"))
                            {
                                string[] sub2 = lne.Split(' ');


                                erd.planID = sub2[1];
                            }
                    }
                    catch
                    {
                        MessageBox.Show("Erreur lors de la lecture du planID. Veuillez vérifier le format du rapport : " + lne);
                    }
                    #endregion



                }

                if (erd.patientID != ctx.Patient.Id)
                {
                    MessageBox.Show("Un document de type Dosimétrie a été trouvé mais le Patient ID du rapport PDF ne correspond pas au Patient ID du plan en cours.\n\n" +
                        "Patient ID du plan en cours: " + ctx.Patient.Id + "\n" +
                        "Patient ID du rapport PDF: " + erd.patientID);

                }

            }
            */


        }

        static string readWithPdfPig(string pathToPdf)
        {
            var sb = new StringBuilder();

            using (UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(pathToPdf))
            {
                int nbPages = document.NumberOfPages;
                for (int i = 1; i <= nbPages; i++)
                {
                    Page page = document.GetPage(i);

                    sb.AppendLine($"----- Page {i}/{nbPages} -----");

                    // Regroupe les mots en lignes selon leur coordonnée verticale (Y)
                    var lignes = ReconstituerLignes(page);

                    foreach (string ligne in lignes)
                        sb.AppendLine(ligne);

                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }

        static List<string> ReconstituerLignes(Page page, double toleranceY = 3.0)
        {
            var lignes = new List<string>();

            // On récupère tous les mots avec leur position Y
            var mots = page.GetWords();

            // On groupe les mots par Y (valeurs proches => même ligne)
            var groupes = mots
                .GroupBy(w => Math.Round(w.BoundingBox.Bottom / toleranceY))
                .OrderByDescending(g => g.Key); // du haut vers le bas

            foreach (var g in groupes)
            {
                var ligne = string.Join(" ", g.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text));
                lignes.Add(ligne);
            }

            return lignes;
        }
        private DateTime convertToDateTime(String dateString) // string to Date
        {
            DateTime dt;

            var cultureInfo = new CultureInfo("fr-FR");
            DateTime.TryParseExact(dateString, "dddd d MMMM yyyy HH:mm:ss", cultureInfo, DateTimeStyles.None, out dt);

            return dt;

        }
        public static bool theseStringsAreEqual(string s1, string s2) // compare no spacpe and no case
        {
            string cleaned1 = new string(s1.Where(c => !char.IsWhiteSpace(c)).ToArray());
            string cleaned2 = new string(s2.Where(c => !char.IsWhiteSpace(c)).ToArray());

            return string.Equals(cleaned1, cleaned2, StringComparison.OrdinalIgnoreCase);
        }
        string getStructureNameInTheLine(string line)
        {
            string s = line;
            // from
            // CanalMed+5 232.47 0.14 2.92 17.10 n/a n/a n/a n/a n/a Exit Only
            // to 
            // CanalMed+5
            try
            {
                int spaceCount = line.Count(c => c == ' '); // count spaces
                s = string.Join(" ", line.Split(' ').Take(spaceCount - 10));
            }
            catch
            {
                s = string.Join(" ", line.Split(' ').Take(1));
            }
            return s;
        }
        bool emptyString(string a)
        {
            bool empty = false;

            if (a == "" || a == null)
                empty = true;

            return empty;

        }
        public void checkTRD(tomoReportData mytrd)
        {

            String msg = String.Empty;


            if (emptyString(trda.planName)) msg += "Le nom du plan n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.planType)) msg += "Le type de plan n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.machineNumber)) msg += "Le numéro de machine n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.machineRevision)) msg += "Le numéro de révision n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.prescriptionMode)) msg += "Le mode de prescription n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.prescriptionTotalDose == 0) msg += "La prescription totale n'a pas été trouvée dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.prescriptionStructure)) msg += "Le structure de prescription n'a pas été trouvée dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.prescriptionMode)) msg += "Le mode de prescriptionn n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.prescriptionDosePerFraction == 0) msg += "La doe par fraction n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.prescriptionNumberOfFraction == 0) msg += "Le nombre de fractions  n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.maxDose == 0) msg += "La dose max. n'a pas été trouvée dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.approvalStatus)) msg += "Le statut de plan n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.MUplanned == 0) msg += "Le nombre d'UM n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.MUplannedPerFraction == 0) msg += "Le nombre d'UM par fraction n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.fieldWidth == 0) msg += "La largeur du fx n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.pitch == 0) msg += "Le pitch du plan n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.modulationFactor == 0) msg += "Le facteur de modulation n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.gantryPeriod == 0) msg += "La période de gantry n'a pas été trouvée dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.couchTravel == 0) msg += "Le mouvement de la table n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.couchSpeed == 0) msg += "La vitesse de table n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.beamOnTime == 0) msg += "beamOnTime n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.deliveryMode)) msg += "deliveryMode n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.resolutionCalculation)) msg += "resolutionCalculation n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (trda.refDose == 0) msg += "refDose n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.patientName)) msg += "patientName n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.patientID)) msg += "patientID n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.HUcurve)) msg += "HUcurve n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.approverID)) msg += "La signature médicale n'a pas été trouvée dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.patientPosition)) msg += "patientPosition n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            //if (trd.originX == 99999) msg += "originX n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.planningMethod)) msg += "planningMethod n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";
            if (emptyString(trda.CTDate)) msg += "CTDate n'a pas été trouvé dans le rapport pdf de Dosimétrie Tomo\n";


            if (msg != "")
            {
                msg += "\n\nVérifiez les options d'impression de Precision. Il faut cocher DICOM SERIES et DENSITY MODEL";
                System.Windows.MessageBox.Show(msg);
            }


        }
        public void displayInfo()
        {
            String s = null;
            s += " Plan name: " + trda.planName;
            s += "\n Plan type: " + trda.planType;
            s += "\n MachineID: " + trda.machineNumber;
            s += "\n Machine rev: " + trda.machineRevision;
            s += "\n prescriptionMode: " + trda.prescriptionMode;
            s += "\n prescriptionTotalDose: " + trda.prescriptionTotalDose;
            s += "\n prescriptionStructure: " + trda.prescriptionStructure;

            s += "\n prescriptionDosePerFraction: " + trda.prescriptionDosePerFraction;
            s += "\n prescriptionNumberOfFraction: " + trda.prescriptionNumberOfFraction;
            s += "\n approvalStatus: " + trda.approvalStatus;
            s += "\n MUplanned: " + trda.MUplanned;
            s += "\n MUplannedPerFraction: " + trda.MUplannedPerFraction;
            System.Windows.MessageBox.Show(s);

            s = null;
            s += "\n Field Width: " + trda.fieldWidth;

            s += "\n isDynamic: " + trda.isDynamic;

            s += "\n pitch: " + trda.pitch;
            s += "\n modulationFactor: " + trda.modulationFactor;
            s += "\n gantryNumberOfRotation: " + trda.gantryNumberOfRotation;

            s += "\n gantryPeriod: " + trda.gantryPeriod;
            s += "\n couchTravel: " + trda.couchTravel;
            s += "\n couchSpeed: " + trda.couchSpeed;
            s += "\n redLaserXoffset: " + trda.redLaserXoffset;
            s += "\n redLaserYoffset: " + trda.redLaserYoffset;
            s += "\n redLaserZoffset: " + trda.redLaserZoffset;
            s += "\n beamOnTime: " + trda.beamOnTime;

            s += "\n deliveryMode: " + trda.deliveryMode;
            s += "\n algorithm: " + trda.algorithm;
            s += "\n resolutionCalculation: " + trda.resolutionCalculation;
            s += "\n refDose: " + trda.refDose;
            s += "\n refPointX: " + trda.refPointX;
            s += "\n refPointY: " + trda.refPointY;
            s += "\n refPointZ: " + trda.refPointZ;

            s += "\n planningMethod: " + trda.planningMethod;
            s += "\n patientName: " + trda.patientName;
            s += "\n patientFirstName: " + trda.patientFirstName;
            s += "\n patientID: " + trda.patientID;


            s += "\n HUcurve: " + trda.HUcurve;
            s += "\n approverID: " + trda.approverID;
            s += "\n patientPosition: " + trda.patientPosition;
            s += "\n originX: " + trda.originX;
            s += "\n originY: " + trda.originY;
            s += "\n originZ: " + trda.originZ;





            System.Windows.MessageBox.Show(s);
        }
        public tomoReportData Trd { get => trda; set => trda = value; }
        public eclipseReportData Erd { get => erd; set => erd = value; }

    }
    public class tomoReportData
    {
        public string planReportDataAsATextFilePath;
        public string planName;
        public string planType;
        public string machineNumber;
        public string machineRevision;
        public string prescriptionMode;
        public string prescriptionStructure;
        public double prescriptionTotalDose;
        public double prescriptionDosePerFraction;
        public int prescriptionNumberOfFraction;
        public string approvalStatus;
        public double MUplanned;
        public double MUplannedPerFraction;
        public double fieldWidth;
        public bool isDynamic;
        public double pitch;
        public double modulationFactor;
        public double gantryNumberOfRotation;
        public double gantryPeriod;
        public double couchTravel;
        public double couchSpeed;
        public double redLaserXoffset;
        public double redLaserYoffset;
        public double redLaserZoffset;
        public double beamOnTime;
        public string deliveryMode;
        public string algorithm;
        public string resolutionCalculation;
        public double refDose;
        public double refPointX;
        public double refPointY;
        public double refPointZ;
        public string planningMethod;
        public string patientName;
        public string patientFirstName;
        public string patientID;
        //public String patientDateOfBirth;
        //public String patientSex;
        public string HUcurve;
        public string approverID;
        public string approvalDate;
        //public bool isHeadFirst;
        //public bool isSupine;
        public string patientPosition;
        public double originX;
        public double originY;
        public double originZ;
        public double maxDose;
        public string CTDate;
        public List<string> blockedOAR;// = new List<String>();
                                       //public int numberOfCTslices;



        public tomoReportData()  //Constructor. 
        {
            blockedOAR = new List<string>();

        }
        public void deleteTomoReportData()
        {
            if (System.IO.File.Exists(planReportDataAsATextFilePath))
            {
                System.IO.File.Delete(planReportDataAsATextFilePath);
            }
        }
    }
    public class eclipseReportData
    {
        public string planID;//{ get; set; } //Plan ID
        public DateTime approDate;// { get; set; } //Approval Date
        public string patientID;//{ get; set; } //Patient ID
        public string courseID;//{ get; set; } //Course ID




        public eclipseReportData()  //Constructor. 
        {


        }
    }


}
