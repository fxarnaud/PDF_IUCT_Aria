using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using PdfSharp.Pdf;
using CommunityToolkit.Mvvm;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using VMS.TPS.Common.Model.API;
using VMS.OIS.ARIALocal.WebServices.Document.Contracts;
using PdfSharp.Pdf.IO;
using MigraDoc.DocumentObjectModel;

namespace PDF_IUCT
{
    public class AriaSender
    {

        ScriptContext _ctx;
        string _filepath;
        private byte[] _binaryContent;
        private string _patientId;
        private User _appUser;
        private string _templateName;
        private DocumentType _documentType;

        public AriaSender(ScriptContext ctx, string path)
        {
            _ctx = ctx;
            _filepath = path;
            GetDocInfo();
            SendToAria();
        }

        private void GetDocInfo()
        {
            _patientId = _ctx.Course.Patient.Id;
            _appUser = _ctx.CurrentUser;
            _templateName = _ctx.ExternalPlanSetup.Id;
            _documentType = new DocumentType
            {
                DocumentTypeDescription = "Dosimétrie"
            };
        }

        private void SendToAria()
        {
            //Recuperation du pdf et passage en binaire
            PdfDocument doc = PdfReader.Open(_filepath);

            //outputDocument
            MemoryStream stream = new MemoryStream();
            //klskd
            doc.Save(stream, false);

            //outputDocument.Save(stream, false);
            _binaryContent = stream.ToArray();

            //Creation du document a envoyer, recuperaion des infos et envoi vers aria
            CustomInsertDocumentsParameter.PostDocumentData(_patientId, _appUser,
                _binaryContent, _templateName, _documentType);
        }


    }
}
