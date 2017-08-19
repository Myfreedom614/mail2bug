using System.Collections.Generic;
using System.Linq;
using log4net;
using Mail2Bug.ExceptionClasses;
using Microsoft.Exchange.WebServices.Data;
using System;

namespace Mail2Bug.Email.EWS
{
    /// <summary>
    /// This implementation of IMailboxManager monitors a specific folder belonging to an exchange
    /// user. All messages coming into the folder will be retrieved via ReadMessages
    /// </summary>
    public class FolderMailboxManager : IMailboxManager
    {
        private readonly ExchangeService _service;
        private readonly string _mailFolder;
        private readonly IMessagePostProcessor _postProcessor;
        private readonly bool _useConversationGuidOnly;
        IEnumerable<string> _recipients;

        public FolderMailboxManager(ExchangeService connection, string incomingFolder, IMessagePostProcessor postProcessor, bool useConversationGuidOnly, IEnumerable<string> recipients = null)
        {
            _service = connection;
            _mailFolder = incomingFolder;
            _postProcessor = postProcessor;
            _useConversationGuidOnly = useConversationGuidOnly;
            _recipients = recipients;
        }

        public IEnumerable<IIncomingEmailMessage> ReadMessages()
        {
            var folder = FolderNameResolver.FindFolderByName(_mailFolder, _service);

            if (folder == null)
            {
                Logger.ErrorFormat("Couldn't find incoming mail folder ({0})", _mailFolder);
                throw new MailFolderNotFoundException(_mailFolder);
            }

            if (folder.TotalCount == 0)
            {
                Logger.DebugFormat("No items found in folder '{0}'. Returning empty list.", _mailFolder);
                return new List<IIncomingEmailMessage>();
            }

            var items = folder.FindItems(new ItemView(folder.TotalCount)).OrderBy(x => x.DateTimeReceived);

            if (_recipients == null)
                return items
                .Where(item => item is EmailMessage)
                .OrderBy(message => message.DateTimeReceived)
                .Select(message => new EWSIncomingMessage(message as EmailMessage, this._useConversationGuidOnly))
                .AsEnumerable();
            else
                return items
                .Where(item => item is EmailMessage && ShouldConsiderMessage(item as EmailMessage, _recipients.ToArray()))
                .OrderBy(message => message.DateTimeReceived)
                .Select(message => new EWSIncomingMessage(message as EmailMessage, this._useConversationGuidOnly))
                .AsEnumerable();
        }

        public void OnProcessingFinished(IIncomingEmailMessage message, bool successful)
        {
            _postProcessor.Process((EWSIncomingMessage)message, successful);
        }

        private static bool ShouldConsiderMessage(EmailMessage message, string[] recipients)
        {
            if (message == null)
            {
                return false;
            }

            // Load additional properties for the current EmailMessage item
            message.Load();

            // If no recipients were mentioned, it means process all incoming emails
            if (!recipients.Any())
            {
                return true;
            }

            var msg = GetToCCList(message);

            // If the recipient is in either the To or CC lines, then this message should be considered
            return recipients.Any(recipient =>
                EmailAddressesMatch(msg.ToAddresses, recipient) ||
                EmailAddressesMatch(msg.ToNames, recipient) ||
                EmailAddressesMatch(msg.CcAddresses, recipient) ||
                EmailAddressesMatch(msg.CcNames, recipient));
        }

        private static MessageToCC GetToCCList(EmailMessage message)
        {
            var m = new MessageToCC();
            var addlist = new List<string>();
            var namelist = new List<string>();
            foreach (EmailAddress msg in message.ToRecipients)
            {
                addlist.Add(msg.Address);
                namelist.Add(msg.Name);
            }
            m.ToAddresses = addlist;
            m.ToNames = namelist;

            var ccAddlist = new List<string>();
            var ccNamelist = new List<string>();
            foreach (EmailAddress msg in message.CcRecipients)
            {
                ccAddlist.Add(msg.Address);
                ccNamelist.Add(msg.Name);
            }
            m.CcAddresses = ccAddlist;
            m.CcNames = ccNamelist;

            return m;
        }

        private static bool EmailAddressesMatch(IEnumerable<string> emailAddresses, string recipient)
        {
            return emailAddresses != null &&
                emailAddresses.Any(address =>
                    address != null && address.Equals(recipient, StringComparison.InvariantCultureIgnoreCase));
        }

        private class MessageToCC
        {
            public IEnumerable<string> ToAddresses { get; set; }
            public IEnumerable<string> CcAddresses { get; set; }
            public IEnumerable<string> ToNames { get; set; }
            public IEnumerable<string> CcNames { get; set; }
        }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(FolderMailboxManager));
    }
}
