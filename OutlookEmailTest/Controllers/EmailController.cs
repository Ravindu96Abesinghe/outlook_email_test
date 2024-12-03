using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailTest.Controllers
{
    public class EmailController : Controller
    {
        public ActionResult OpenOutlook()
        {
            try
            {
                // Create an Outlook application instance
                var outlookApp = new Outlook.Application();

                // Create a new email item
                var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                // Set email properties
                mailItem.To = "recipient@example.com";
                mailItem.Subject = "Test Email";
                mailItem.Body = "This is a test email from the MVC project.";

                //// Optionally attach a file
                //string pdfPath = Server.MapPath("~/Content/Sample.pdf"); // Example path
                //if (System.IO.File.Exists(pdfPath))
                //{
                //    mailItem.Attachments.Add(pdfPath, Outlook.OlAttachmentType.olByValue, 1, "Sample.pdf");
                //}

                // Display the email in Outlook
                mailItem.Display();

                return Content("Outlook email opened successfully!");
            }
            catch (System.Exception ex)
            {
                return Content($"Error: {ex.Message}");
            }
        }
    }
}
