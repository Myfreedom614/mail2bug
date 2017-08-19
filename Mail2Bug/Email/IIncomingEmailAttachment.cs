using System.Collections.Generic;

namespace Mail2Bug.Email
{
    /// <summary>
    /// Similar to IIncomingEmailMessage, this interface is an adapter for attachments of incoming email messages.
    /// </summary>
    public interface IIncomingEmailAttachment
    {
        string SaveAttachmentToFile(List<string> attNameList = null);
        string SaveAttachmentToFile(string filename);
    }
}
