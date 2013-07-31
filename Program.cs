using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Security;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;

using Microsoft.Exchange.WebServices.Data;

using System.Data;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Drawing; 


namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            service.Credentials = new WebCredentials("ia32el", "Bits@123");

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;

            service.AutodiscoverUrl("cynthia.w.li@intel.com", RedirectionUrlValidationCallback);
            
            DisplayInbox(service);
            // ToDB(service);
            Console.WriteLine("\r\n");
            Console.WriteLine("Press any key...");

            Console.Read();
        }

        // validates whether the redirected URL is using Transport Layer Security.
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }


        private static bool CertificateValidationCallBack(
                            object sender,
                            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                            System.Security.Cryptography.X509Certificates.X509Chain chain,
                            System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }
        public class MailItem
        {
            public string From;
            public string id;
            public string Subject;
            public string Time;
            public string Body;
            public string category;
        }

        public static string findcategory(string subject) {
            if (subject.IndexOf("bug") != -1 || subject.IndexOf("Bug") != 1) return "Bug";
            else if (subject.IndexOf("RE:") != -1 && subject.IndexOf("CI") != -1) return "CI";
            else if (subject.IndexOf("Houdini") != -1 && subject.IndexOf("ci") == -1) return "Houdini";
            else if (subject.IndexOf("app") != -1 || subject.IndexOf("App") != -1) return "Application";
            else if (subject.IndexOf("report") != -1 || subject.IndexOf("Report") != -1) return "Report";
            else if (subject.IndexOf("test") != -1 || subject.IndexOf("Test") != -1) return "Test";
            else return "Others";
        }
        public static String format_body(String s) {
            String body = "";
            int a = s.IndexOf("<body");
            int b = s.IndexOf("From:</span>");
            if (a != -1)
            {
                if (b != -1)
                {
                    body += s.Substring(a, b - a);
                    body += "</body>";
                }
                else body += s.Substring(a);
                s = body;
            }
            /*
            s = s.Replace("&", "&amp;");
            s = s.Replace(">", "&gt;");
            s = s.Replace("<", "&lt;");
            s = s.Replace("'", "&apos;");
            s = s.Replace("\"", "&quot;");
            */
            s = s.Replace("&", "%26");
            s = s.Replace("%", "%25");
            return s;
        }
        public static MailItem[] GetUnreadMailFromInbox(ExchangeService service)
        {
            ItemView view = new ItemView(128);
            view.OrderBy.Add(ItemSchema.DateTimeCreated, Microsoft.Exchange.WebServices.Data.SortDirection.Descending);

            SearchFilter.SearchFilterCollection searchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);
            // Add the search filter collection to filter the emails that contain any keyword in the contain file.
            string[] lines = System.IO.File.ReadAllLines("containkeywords.txt");
            foreach (string line in lines)
            {
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, line));
                // Add the search filter
            }
            SearchFilter.SearchFilterCollection notsearchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
            // Add the search filter collection to filter the eamils that cannot contain any keywords in the notcontain file
            lines = System.IO.File.ReadAllLines("notcontainkeywords.txt");
            foreach (string line in lines)
            {
                notsearchFilterCollection.Add(new SearchFilter.Not(new SearchFilter.ContainsSubstring(ItemSchema.Subject, line)));
            }
            SearchFilter.SearchFilterCollection sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
            sf.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            sf.Add(searchFilterCollection);
            sf.Add(notsearchFilterCollection);

            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);
            if (findResults.TotalCount > 0)
            {
                ServiceResponseCollection<GetItemResponse> items =
                    service.BindToItems(findResults.Select(item => item.Id), new PropertySet(BasePropertySet.FirstClassProperties,
                                        EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeCreated));
                return items.Select(item =>
                {
                    return new MailItem()
                    {
                        From = ((Microsoft.Exchange.WebServices.Data.EmailAddress)item.Item[EmailMessageSchema.From]).Address,
                        id = item.Item.Id.ToString(),
                        Subject = item.Item.Subject.Replace("&", "%26").Replace("%", "%25"),
                        Time = item.Item.DateTimeCreated.ToString(),
                        Body = format_body(item.Item.Body.ToString()),
                        category = findcategory(item.Item.Subject),
                    };
                }).ToArray();
            }
            else return null;
        }

        public static void DisplayInbox(ExchangeService service)
        {

            StringBuilder sb = new StringBuilder();
            StringWriter Response = new StringWriter(sb);
            //System.IO.StreamWriter Response = new System.IO.StreamWriter("emails.xml");

            #region ReadEmail
            

            Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox); // Inbox folder

            MailItem[] unreadMails = new MailItem[400];
            unreadMails = GetUnreadMailFromInbox(service);

            Response.Write("source=email&content=");
            Response.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            Response.WriteLine("<items>");

            if (unreadMails != null)
            {
                foreach (MailItem em in unreadMails)
                {
                    
                    Response.WriteLine("<item>");
                    Response.WriteLine("<idemails>NULL</idemails>");
                    Response.WriteLine("<category>" + em.category + "</category>");
                    Response.WriteLine("<subject>" + em.Subject + "</subject>");
                    Response.WriteLine("<from_add>" + em.From + "</from_add>");
                    Response.WriteLine("<body><![CDATA[" + em.Body + "]]></body>");
                    //Response.WriteLine("<body>" + em.Body + "</body>");
                    Response.WriteLine("<time>" + em.Time + "</time>");
                    Response.WriteLine("<is_insert>true</is_insert>");
                    Response.WriteLine("</item>");
                }
            }
            Response.WriteLine("</items>");

            #endregion

            
            // Create POST data and convert it to a byte array. 
            String postData = sb.ToString();
            // Create a request using a URL that can receive a post.
            String url = "http://taotieserver.sh.intel.com:8080/servlet";
            WebRequest request = WebRequest.Create(url);
            // Set the Method property of the request to POST.
            request.Method = "POST";
            Console.WriteLine(postData);
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            
            // Set the ContentType property of the WebRequest.
            request.ContentType = "application/x-www-form-urlencoded";
            // Set the ContentLength property of the WebRequest.
            request.ContentLength = byteArray.Length;

            // Get the request stream.
            Stream dataStream = request.GetRequestStream();
            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);
            // Close the Stream object.
            dataStream.Close();
            // Get the response.
            WebResponse response = request.GetResponse();
            // Display the status.
            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
            // Get the stream containing content returned by the server.
            dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            String responseFromServer = reader.ReadToEnd();
            // Display the content.
            Console.WriteLine(responseFromServer);
            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();
            
            Response.Close();
        }


        public static void ToDB(ExchangeService service)
        {
            MailItem[] unreadMails = new MailItem[100];
            unreadMails = GetUnreadMailFromInbox(service);

            String db = "server=10.239.51.65; database=emails; userid=root; password=intel123;";
            //String db = "server=10.239.51.53; database=emails; userid=root; password=1;";
            MySqlConnection con = null;
            try
            {
                con = new MySqlConnection(db);
                con.Open(); 
                if (unreadMails != null)
                {
                    foreach (MailItem em in unreadMails)
                    {
                        if (em.category == "CI")
                        {
                            String cmdText = "INSERT INTO email(id, subject, from_add, body, time) VALUES(NULL, @suject, @from, @body, @time)";
                            MySqlCommand cmd = new MySqlCommand(cmdText, con);
                            cmd.Prepare();
                            //we will bound a value to the placeholder
                            cmd.Parameters.AddWithValue("@subject", em.Subject);
                            cmd.Parameters.AddWithValue("@from", em.From);
                            cmd.Parameters.AddWithValue("@body", em.Body);
                            cmd.Parameters.AddWithValue("@time", em.Time);
                            cmd.ExecuteNonQuery(); //execute the mysql command
                        }
                    }
                }
            }
            catch (MySqlException err)
            {
                Console.WriteLine("Error: " + err.ToString());
            }
            finally
            {
                if (con != null)
                {
                    con.Close(); 
                }
            }
        }
    }
}
