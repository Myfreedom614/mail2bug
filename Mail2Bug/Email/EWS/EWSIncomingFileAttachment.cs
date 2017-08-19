using System.IO;
using log4net;
using Mail2Bug.Helpers;
using System.Collections.Generic;

namespace Mail2Bug.Email.EWS
{
    class EWSIncomingFileAttachment : IIncomingEmailAttachment
    {
        private readonly Microsoft.Exchange.WebServices.Data.FileAttachment _attachment;
        private string _newName = null;

        public EWSIncomingFileAttachment(Microsoft.Exchange.WebServices.Data.FileAttachment attachment, string newName=null)
        {
            _attachment = attachment;
            if (newName != null)
                _newName = newName;
        }

        public string SaveAttachmentToFile(List<string> attNameList = null)
        {
            string fileName = _newName == null ? _attachment.Name : _newName;
            var baseFilename = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            if(attNameList != null && attNameList.Contains(baseFilename))
            {
                return null;
            }
            return SaveAttachmentToFile(FileUtils.GetValidFileName(baseFilename, extension, Path.GetTempPath()));
        }

        public string SaveAttachmentToFile(string filename)
        {
            if(!FileUtils.CheckIfExistLastChar(Path.GetFileNameWithoutExtension(filename), '_')) {
                Logger.DebugFormat("Saving attachment named '{0}' to file {1}", _attachment.Name ?? "", filename);

                _attachment.Load(filename);

                return filename;
            }
            return null;
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(EWSIncomingFileAttachment));
    }
}
