using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using VMS.OIS.ARIAExternal.WebServices.Documents.Contracts;
using VMS.TPS.Common.Model.API;

namespace catchDose
{
    internal class docFinder
    {
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

     
            var VisitNoteList = new List<string>();
            int visitnoteloc = response.IndexOf("PtVisitNoteId");
            while (visitnoteloc > 0)
            {
                VisitNoteList.Add(response.Substring(visitnoteloc + 15, 2).Replace(",", ""));
                visitnoteloc = response.IndexOf("PtVisitNoteId", visitnoteloc + 1);
            }
            var response_Doc = JsonConvert.DeserializeObject<DocumentsResponse>(response); // get the list of documents
            var DocTypeList = new List<string>();
            var docNameList = new List<string>();
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
                    
                    docExtractor docExtractor = new docExtractor(response_docdetails, thisDocType, thisDocName, ctx, indexer);

                   if (thisDocType.Contains("Dosimétrie") || thisDocType.Contains("Fiche de positionnement"))
                        docNameList.Add(thisDocName);

                    DateServiceList.Add(dtDateTime);
                    DocTypeList.Add(thisDocType);
                    DocIndexList.Add(loopnum);
                    
                    
                }
                #endregion


                loopnum++;
            }
            #endregion

            String outmsg = "Ces documents ont été extraits d'Aria Documents:\n\n";
            foreach(string s in docNameList)
            {
                outmsg += " - " +  s + "\n";
            }
           
            System.Windows.MessageBox.Show(outmsg);





        }

    }

    public class docExtractor
    {
        private tomoReportData trda;
        private eclipseReportData erd;

        private string outpath = "";

      
        public docExtractor(string response_details, string doctype, string docname, ScriptContext _ctx, int indexer)
        {
            bool typeIsKnown = true;
            string extension = ".pdf";
            if (doctype.Contains("Dosimétrie"))
                extension = ".pdf";
            else if (doctype.Contains("Fiche de positionnement"))
                extension = ".docx";
            else
                typeIsKnown = false;

            if (typeIsKnown)
            {
               
                #region Convert response details (string) to a temp PDF file
                String saveFilePathDir = @"\\srv015\sf_com\simon_lu\milady";
                saveFilePathDir += @"\" + _ctx.Patient.Id + "_" + _ctx.Patient.LastName;

                if (!Directory.Exists(saveFilePathDir))
                    Directory.CreateDirectory(saveFilePathDir);

                string cleanType = doctype.Trim().Replace(" ", "");
                string cleanName = docname.Trim().Replace(" ", "");
                String saveFilePathTemp = saveFilePathDir + @"\" + indexer + "_" + cleanType + "_" + cleanName + extension;
                String saveFilePathTempresponse = saveFilePathDir + @"\" + indexer + "__temp__.txt";

                int startBinary = response_details.IndexOf("\"BinaryContent\"") + 17;
                int endBinary = response_details.IndexOf("\"Certifier\"") - 2;
                string binaryContent2 = response_details.Substring(startBinary, endBinary - startBinary);
                binaryContent2 = binaryContent2.Replace("\\", "");  // the \  makes the string a non valid base64 string                       
                File.WriteAllBytes(saveFilePathTemp, Convert.FromBase64String(binaryContent2));

                #endregion

            }




        }
        
        bool emptyString(string a)
        {
            bool empty = false;

            if (a == "" || a == null)
                empty = true;

            return empty;

        }
  
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
