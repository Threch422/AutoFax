using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FAXCOMEXLib;

namespace AutoFax
{
    // With Reference on https://www.codeproject.com/Articles/1159834/Send-Fax-with-fax-modem-in-Csharp
    public class CustomFaxSender
    {
        private static FaxServer faxServer;
        public CustomFaxSender()
        {
            try
            {
                faxServer = new FaxServer();
                faxServer.Connect("");
                RegisterFaxServerEvents();
                Console.WriteLine("Connected to Fax Service.");
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
        }

        #region Event Handlers/Listeners
        private static void faxServer_OnOutgoingJobAdded(FaxServer pFaxServer, string bstrJobId)
        {
            Console.WriteLine("OnOutgoingJobAdded event fired. A fax is added to the outgoing queue.");
        }

        private static void faxServer_OnOutgoingJobChanged(FaxServer pFaxServer, string bstrJobId, FaxJobStatus pJobStatus)
        {
            Console.WriteLine("OnOutgoingJobChanged event fired. A fax is changed to the outgoing queue.");
            pFaxServer.Folders.OutgoingQueue.Refresh();
            PrintFaxStatus(pJobStatus);
        }

        private static void faxServer_OnOutgoingJobRemoved(FaxServer pFaxServer, string bstrJobId)
        {
            Console.WriteLine("OnOutgoingJobRemoved event fired. Fax job is removed to outbound queue.");
        }
        #endregion

        private static void PrintFaxStatus(FaxJobStatus faxJobStatus)
        {
            if (faxJobStatus.ExtendedStatusCode == FAX_JOB_EXTENDED_STATUS_ENUM.fjesDIALING)
            {
                Console.WriteLine("Dialing...");
            }

            if (faxJobStatus.ExtendedStatusCode == FAX_JOB_EXTENDED_STATUS_ENUM.fjesTRANSMITTING)
            {
                Console.WriteLine("Sending Fax...");
            }

            if (faxJobStatus.Status == FAX_JOB_STATUS_ENUM.fjsCOMPLETED
                && faxJobStatus.ExtendedStatusCode == FAX_JOB_EXTENDED_STATUS_ENUM.fjesCALL_COMPLETED)
            {
                Console.WriteLine("Fax is sent successfully.");
            }
        }

        protected virtual void RegisterFaxServerEvents()
        {
            faxServer.OnOutgoingJobAdded += new IFaxServerNotify2_OnOutgoingJobAddedEventHandler(faxServer_OnOutgoingJobAdded);
            faxServer.OnOutgoingJobChanged +=  new IFaxServerNotify2_OnOutgoingJobChangedEventHandler(faxServer_OnOutgoingJobChanged);
            faxServer.OnOutgoingJobRemoved +=  new IFaxServerNotify2_OnOutgoingJobRemovedEventHandler(faxServer_OnOutgoingJobRemoved);

            var eventsToListen =
                      FAX_SERVER_EVENTS_TYPE_ENUM.fsetFXSSVC_ENDED | FAX_SERVER_EVENTS_TYPE_ENUM.fsetOUT_QUEUE
                    | FAX_SERVER_EVENTS_TYPE_ENUM.fsetOUT_ARCHIVE | FAX_SERVER_EVENTS_TYPE_ENUM.fsetQUEUE_STATE
                    | FAX_SERVER_EVENTS_TYPE_ENUM.fsetACTIVITY | FAX_SERVER_EVENTS_TYPE_ENUM.fsetDEVICE_STATUS;

            faxServer.ListenToServerEvents(eventsToListen);
        }

        /// <summary>
        /// Setup the FaxDocument
        /// </summary>
        /// <param name="recipientInfos">Key: FaxNumber, Value: Recipient Name</param>
        /// <param name="body"></param>
        /// <param name="subject"></param>
        /// <param name="docName"></param>
        /// <returns>
        /// Return the set up FaxDocument
        /// </returns>
        protected virtual FaxDocument FaxDocSetup(Dictionary<string, string> recipientInfos, string body, string subject = null, string docName = null)
        {
            var faxDoc = new FaxDocument
            {
                Priority = FAX_PRIORITY_TYPE_ENUM.fptHIGH,
                ReceiptType = FAX_RECEIPT_TYPE_ENUM.frtNONE,
                AttachFaxToReceipt = true,
                Subject = subject ?? string.Empty,
                Body = body ?? string.Empty,
                DocumentName = docName ?? string.Empty,
            };

            foreach (var recipientInfo in recipientInfos)
            {
                faxDoc.Recipients.Add(recipientInfo.Key, recipientInfo.Value);
            }

            return faxDoc;
        }

        /// <summary>
        /// Send the set up FaxDocument to the specific recipients
        /// The body parameter is expected to be the path of attactment
        /// </summary>
        protected internal virtual void SendFax(Dictionary<string, string> recipientInfos, string body, string subject = null, string docName = null)
        {
            try
            {
                var faxDoc = this.FaxDocSetup(recipientInfos, body, subject, docName);
                var faxRtnValue = faxDoc.Submit(faxServer.ServerName);
                faxDoc = null;

                Console.WriteLine($"Document {Path.GetFileName(body)} has been submitted.");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error in sending fax { Path.GetFileName(body) } to fax server. Error Message: {e.Message}");
            }
        }

        protected internal virtual void DisconnectFaxServer()
        {
            faxServer.Disconnect();
            Console.WriteLine("Disconnected with the Fax Service.");
        }
    }
}
