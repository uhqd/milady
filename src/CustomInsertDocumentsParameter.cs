using Newtonsoft.Json;
using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using VMS.OIS.ARIALocal.WebServices.Document.Contracts;
namespace catchDose
{
    public class CustomInsertDocumentsParameter
    {
        public PatientIdentifier PatientId { get; set; }
        public string DateOfService { get; set; }
        public string DateEntered { get; set; }
        public string BinaryContent { get; set; }
        public DocumentUser AuthoredByUser { get; set; }
        public DocumentUser SupervisedByUser { get; set; }
        public DocumentUser EnteredByUser { get; set; }
        public FileFormat FileFormat { get; set; }
        public DocumentType DocumentType { get; set; }
        public string TemplateName { get; set; }
        public bool IsApproved { get; set; } = true;
        public DocumentUser ApprovedByUser { get; set; }
        public string DateApproved { get; set; }


        public static bool PostDocumentDataPush(string patientId, VMS.TPS.Common.Model.API.User user, byte[] binaryContent, string templateName, DocumentType documentType)
        {
            ServicePointManager.ServerCertificateValidationCallback += (o, c, ch, er) => true;

   
            #region GET APIKEY

            DocSettings docSet = DocSettings.ReadSettings();
            string docKey = docSet.DocKey;
            string hostName = docSet.HostName;
            string port = docSet.Port;

            #endregion


            var documentPushRequest = new CustomInsertDocumentsParameter
            {
                PatientId = new PatientIdentifier { ID1 = patientId },
                DateOfService = $"/Date({Math.Floor((DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds)})/",
                DateEntered = $"/Date({Math.Floor((DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds)})/",
                BinaryContent = Convert.ToBase64String(binaryContent),
                FileFormat = FileFormat.DOCX,
                AuthoredByUser = new DocumentUser
                {
                    SingleUserId = user.Id
                },
                SupervisedByUser = new DocumentUser
                {
                    SingleUserId = user.Id
                },
                EnteredByUser = new DocumentUser
                {
                    SingleUserId = user.Id
                },
                TemplateName = templateName,
                DocumentType = documentType,  // next lines  approve the document
                IsApproved = true,
                ApprovedByUser = new DocumentUser
                {
                    SingleUserId = user.Id
                },
                DateApproved = $"/Date({Math.Floor((DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds)})/"
            };
            var request_base = "{\"__type\":\"";
            var request_document = $"{request_base}InsertDocumentRequest:http://services.varian.com/Patient/Documents\",{JsonConvert.SerializeObject(documentPushRequest).TrimStart('{')}}}";
            string response_document = SendData(request_document, true, docKey, hostName, port);
            // MessageBox.Show(response_document);
            if (!response_document.Contains("GatewayError"))
            {
                VMS.OIS.ARIAExternal.WebServices.Documents.Contracts.DocumentResponse documentResponse = JsonConvert.DeserializeObject<VMS.OIS.ARIAExternal.WebServices.Documents.Contracts.DocumentResponse>(response_document);
                if (documentResponse != null)
                {
                    // if (documentResponse.PtVisitId != null)
                    //{
                    return true;
                    //}
                }
            }
            return false;
        }



        public static bool PostDocumentData(string patientId, VMS.TPS.Common.Model.API.User user, byte[] binaryContent, string templateName, DocumentType documentType, DocSettings docSet)
        {
            ServicePointManager.ServerCertificateValidationCallback += (o, c, ch, er) => true;
            string docKey = docSet.DocKey;
            string hostName = docSet.HostName;
            string port = docSet.Port;
            /*
            DocSettings.ReadSettings();
            string docKey = DocSettings.DocKey;
            string hostName = DocSettings.HostName;
            string port = DocSettings.Port;
            */


            var documentPushRequest = new CustomInsertDocumentsParameter
            {
                PatientId = new PatientIdentifier { ID1 = patientId },
                DateOfService = $"/Date({Math.Floor((DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds)})/",
                DateEntered = $"/Date({Math.Floor((DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds)})/",
                BinaryContent = Convert.ToBase64String(binaryContent),
                FileFormat = FileFormat.PDF,
                AuthoredByUser = new DocumentUser
                {
                    SingleUserId = user.Id
                },
                SupervisedByUser = new DocumentUser
                {
                    SingleUserId = user.Id
                },
                EnteredByUser = new DocumentUser
                {
                    SingleUserId = user.Id
                },
                TemplateName = templateName,
                DocumentType = documentType
            };
            var request_base = "{\"__type\":\"";
            var request_document = $"{request_base}InsertDocumentRequest:http://services.varian.com/Patient/Documents\",{JsonConvert.SerializeObject(documentPushRequest).TrimStart('{')}}}";
            string response_document = SendData(request_document, true, docKey, hostName, port);
            //MessageBox.Show(response_document);
            if (!response_document.Contains("GatewayError"))
            {
                VMS.OIS.ARIAExternal.WebServices.Documents.Contracts.DocumentResponse documentResponse = JsonConvert.DeserializeObject<VMS.OIS.ARIAExternal.WebServices.Documents.Contracts.DocumentResponse>(response_document);
                if (documentResponse != null)
                {
                    if (documentResponse.PtVisitId != 0)// was null
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        public static string SendData(string request, bool bIsJson, string apiKey, string hostName, string port)
        {
            var sMediaTYpe = bIsJson ? "application/json" :
            "application/xml";
            var sResponse = System.String.Empty;
            using (var c = new HttpClient(new
            HttpClientHandler()
            { UseDefaultCredentials = true, PreAuthenticate = true }))
            {
                if (c.DefaultRequestHeaders.Contains("ApiKey"))
                {
                    c.DefaultRequestHeaders.Remove("ApiKey");
                }
                c.DefaultRequestHeaders.Add("ApiKey", apiKey);
                var gatewayURL = $"https://{hostName}:{port}/Gateway/service.svc/interop/rest/Process";
                var task =
                c.PostAsync(gatewayURL,
                new StringContent(request, Encoding.UTF8,
                sMediaTYpe));
                Task.WaitAll(task);
                var responseTask =
                task.Result.Content.ReadAsStringAsync();
                Task.WaitAll(responseTask);
                sResponse = responseTask.Result;
            }
            return sResponse;
        }
    }
}
