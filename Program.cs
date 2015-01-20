using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Limilabs.Client.IMAP;
using Limilabs.Mail;
using Limilabs.Mail.Headers;
using Limilabs.Mail.MIME;
using ManagedClient;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;

namespace Hybrid
{
    class Program
    {
        public static List<string> emailStuff = new List<string>();
        public static List<string> articlesHTML = new List<string>();
        public static List<string> articlesCorrected = new List<string>();
        public static List<string> articlesHTMLproblems = new List<string>();

        static void Main()
        {
            using (Imap imap = new Imap())
            {
                var mailSettings = ConfigurationManager.GetSection("mailSettings") as NameValueCollection;
                string imapServer = mailSettings["imap"].ToString();
                string smtpServer = mailSettings["smtp"].ToString();
                string mailHost = mailSettings["host"].ToString();
                string mailUser = mailSettings["user"].ToString();
                string mailPass = mailSettings["pass"].ToString();
                string[] reports = mailSettings["reports"].ToString().Split(';');

                //logging the output
                var cc = new ConsoleCopy("log.txt");

                //connect to the server using ssl
                imap.ConnectSSL(imapServer);
                imap.Login(mailUser, mailPass);
                imap.SelectInbox();

                //find all messages
                List<long> uids = imap.Search(Flag.All);
                Console.WriteLine("Number of messages in the Inbox folder: " + uids.Count);

                foreach (long uid in uids)
                {
                    //download and parse each message
                    IMail email = new MailBuilder().CreateFromEml(imap.GetMessageByUID(uid));

                    //display email data, save attachments
                    ProcessMessage(email, uid);
                    imap.MoveByUID(uid, "checked");
                }
                imap.Close();


                if (uids.Count != 0)
                {
                    //forming HTML report
                    emailStuff.Add("<html>");
                    emailStuff.Add("<style>");
                    emailStuff.Add("table.data-table { font-size:13px; border: 0px solid #CCCCBB; margin-bottom: 2em; width: 100%; }");
                    emailStuff.Add("table.data-table th { background: none repeat scroll 0 0 #F0F0F0; border: 1px solid #C9C9C9; color: #555555; text-align: left; }");
                    emailStuff.Add("table.data-table tr { border-bottom: 1px solid #C9C9C9; }");
                    emailStuff.Add("table.data-table td { padding: 10px; }");
                    emailStuff.Add("table.data-table table th { padding: 0px; }");
                    emailStuff.Add("table.data-table td { background: none repeat scroll 0 0 #F6F6F6; border: 1px solid #C9C9C9; }");
                    emailStuff.Add("table.data-table td.orig { background: none repeat scroll 0 0 #c9f6f6; border: 1px solid #C9C9C9; }");
                    emailStuff.Add("table.data-table tr.even td { background: none repeat scroll 0 0 #FCFCFC; }");
                    emailStuff.Add("</style>");

                    emailStuff.Add("<h2>Added articles</h2>");
                    emailStuff.Add("<table class=\"data-table\"><tbody>");
                    emailStuff.Add("<tr><td><b>Number of MEPHI authors</b></td><td><b>Title</b></td><td><b>Journal</b></td>" +
                                     "<td><b>Year</b></td><td><b>DOI</b></td></tr>");
                    foreach (string article in articlesHTML)
                    { emailStuff.Add(article); }
                    emailStuff.Add("</tbody></table>");

                    emailStuff.Add("<h2>Corrected articles:</h2>");
                    emailStuff.Add("<table class=\"data-table\"><tbody>");
                    emailStuff.Add("<tr><td><b>MFN</b></td><td><b>Title</b></td><td><b>DOI</b></td><td><b>Year</b></td></tr>");
                    foreach (string article in articlesCorrected)
                    { emailStuff.Add(article); }
                    emailStuff.Add("</tbody></table>");

                    emailStuff.Add("<h2>Problems:</h2>");
                    emailStuff.Add("<table class=\"data-table\"><tbody>");
                    emailStuff.Add("<tr><td><b>MFN</b></td><td><b>Title</b></td><td><b>DOI</b></td><td><b>Year</b></td></tr>");
                    foreach (string article in articlesHTMLproblems.ToArray())
                    { emailStuff.Add(article); }
                    emailStuff.Add("</tbody></table>");

                    //configuring and sending a report via email
                    Console.WriteLine("--------------");
                    Console.WriteLine("Sending report");

                    //stop logging
                    cc.Dispose();

                    string data = string.Join("", emailStuff.ToArray());
                    SendMessage(data, "mail processing report - new articles", mailUser, mailHost, mailPass, smtpServer, reports, true);

                }
                else
                {
                    emailStuff.Add("<html><b>Inbox folder is empty.</b></html>");
                    string data = string.Join("", emailStuff.ToArray());
                    SendMessage(data, "mail processing report - no new articles", mailUser, mailHost, mailPass, smtpServer, reports, false);
                }
            }
        }

        private static void SendMessage(string data, string subject, string mailUser, string mailHost, string mailPass, string smtpServer, string[] reports, bool log)
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            System.Net.Mail.SmtpClient clientSmtp = new System.Net.Mail.SmtpClient();

            clientSmtp.Credentials = new System.Net.NetworkCredential(mailUser + "@" + mailHost, mailPass);
            clientSmtp.Port = 25;
            clientSmtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
            clientSmtp.EnableSsl = true;
            clientSmtp.Host = smtpServer;

            if (log == true) {mail.Attachments.Add(new System.Net.Mail.Attachment("log.txt")); }
            mail.From = new System.Net.Mail.MailAddress(mailUser + "@" + mailHost);
            foreach (string rep in reports) { mail.To.Add(new System.Net.Mail.MailAddress(rep)); }
            mail.IsBodyHtml = true;
            mail.Subject = subject;
            mail.Body = data;

            clientSmtp.Send(mail);
        }

        private static void ProcessMessage(IMail email, long uid)
        {
            Console.WriteLine("---------------------------");
            Console.WriteLine("#Saving mail [uid={0}]", uid.ToString());
            string fileName = uid.ToString() + ".txt";
            StreamWriter fileTXT = new StreamWriter("./data/" + fileName);

            fileTXT.WriteLine("#Subject: " + email.Subject);
            fileTXT.WriteLine("#From: " + JoinAddresses(email.From));
            fileTXT.WriteLine("#To: " + JoinAddresses(email.To));
            fileTXT.WriteLine("#Date: " + email.Date);
            fileTXT.WriteLine("#Text: " + email.Text);

            string from = JoinAddresses(email.From);
            if (from.Contains("noreply@isiknowledge.com"))
            {
                //Web of Science
                int cWosScopus = 0;
                //get info entry by entry
                string textPattern = @"\*Record [1-9][0-9]* of [1-9][0-9]*(.*?)Cited References";
                MatchCollection articleArray = Regex.Matches(email.Text, textPattern, RegexOptions.Singleline);
                //for each one do stuff with article.Groups[1].Value
                foreach (Match article in articleArray)
                {
                    //dunno if its actually possible to get rid of regex in wos's mail processing
                    string titlePattern = @"\r\nTitle:\r\n(.*?)\r\n\r\n";
                    string authorsPattern = @"\r\nAuthor Full Names:\r\n(.*?)\r\n\r\n";//test
                    string citedPattern = @"\r\nTimes Cited:\r\n(.*?)\r\n";
                    string issnPattern = @"\r\nISSN:\r\n(.*?)\r\n";
                    string urlPattern = @"\*View Full Record:(.*?)\r\n\*Order Full Text";
                    string addPatternAll = @"\r\nSource:\r\n(.*?)\r\n\r\n";
                    string addPatternJV = @"\r\nSource:\r\n(.*?), (.*?)( |:|\()";
                    string addPatternIssue = @",(.*?)\((.*?)\)";
                    string addPatternDOI = @"10.[A-Za-z0-9.\/]*?[\/][A-Za-z0-9\/)(.:_-]*?( |\r\n)";
                    string addPatternPages = @"[0-9_-]*?;";
                    string addPatternYear = @"\s\d{4}\s$";
                    string abstractPattern = @"\r\nAbstract:(.*?)\r\n\r\n";
                    string collabPattern = @"\r\nGroup Author\(s\):\r\n(.*?)\r\n\r\nSource:";
                    string docTypePattern = @"\r\n\r\nDocument Type:\r\n(.*?)\r\n\r\n";

                    Match allAuthors = Regex.Match(article.Groups[1].Value, authorsPattern, RegexOptions.Singleline);
                    Match title = Regex.Match(article.Groups[1].Value, titlePattern);
                    Match cited = Regex.Match(article.Groups[1].Value, citedPattern);
                    Match issn = Regex.Match(article.Groups[1].Value, issnPattern);
                    Match url = Regex.Match(article.Groups[1].Value, urlPattern);
                    Match additional = Regex.Match(article.Groups[1].Value, addPatternAll);
                    Match journal = Regex.Match(article.Groups[1].Value, addPatternJV);
                    Match volume = Regex.Match(article.Groups[1].Value, addPatternJV);
                    Match issue = Regex.Match(additional.Groups[1].Value, addPatternIssue);
                    Match doi = Regex.Match(additional.Groups[1].Value, addPatternDOI);
                    Match pages = Regex.Match(additional.Groups[1].Value, addPatternPages);
                    Match year = Regex.Match(additional.Groups[1].Value, addPatternYear);
                    Match abstr = Regex.Match(article.Groups[1].Value, abstractPattern, RegexOptions.Singleline);
                    Match collab = Regex.Match(article.Groups[1].Value, collabPattern);
                    Match doctype = Regex.Match(article.Groups[1].Value, docTypePattern);

                    System.IO.StreamReader searchFile = new System.IO.StreamReader("searching.ini");
                    String searchStrings = searchFile.ReadToEnd();
                    searchFile.Close();

                    int mephiAuthCount = 0;
                    string[] authorsMEPHI = { };
                    List<string> authorsNOTMEPHI = new List<string>();
                    string[] allAuthorsStr = { }; //all authors
                    string[,] authorsMEPHIArray = new string[100, 100]; // 100 is a max number of (mephi authors + 4 not mephi authors)

                    string wosSearchPattern = @"\[Web of Science\]\r\n(.*?)\r\n\[";
                    Match wosQueries = Regex.Match(searchStrings, wosSearchPattern, RegexOptions.Singleline);
                    string[] charsToSplit = new string[] { "\r\n" };
                    string[] wosArray = wosQueries.Groups[1].Value.Split(charsToSplit, StringSplitOptions.None);

                    string affilPattern = @"\r\nAddresses:\r\n(.*?)\r\n\r\n(.*?):";
                    string affilAuthorsPattern = @"\[(.*?)\](.*?)\.";

                    Match affilWoSAll = Regex.Match(article.Groups[1].Value, affilPattern, RegexOptions.Singleline);
                    Console.WriteLine("-------------------------------------------");
                    Console.WriteLine("------------Affiliation match--------------");
                    string[] affilWosSplit = affilWoSAll.Groups[1].Value.Split(charsToSplit, StringSplitOptions.None);

                    //if there is only one affilation, every single author is from mephi (:
                    if (affilWosSplit.Length == 1)
                    {
                        string tempAffil;
                        Match affilMatch = Regex.Match(affilWoSAll.Groups[1].Value, affilAuthorsPattern, RegexOptions.Singleline);
                        Console.WriteLine("only one affiliation");
                        authorsMEPHI = allAuthors.Groups[1].Value.Replace("\r\n", "").Replace(".", "").Replace(" ", "").Split(';');

                        if (affilMatch.Length != 0) //affiliation is in a normal format: "[Authors]Affiliation"
                        {
                            tempAffil = affilMatch.Groups[2].Value.TrimStart();
                        }
                        else //affiliation is in a plain format: "Affiliation"
                        {
                            tempAffil = affilWoSAll.Groups[1].Value;
                        }

                        //add each author and his afiliation to array
                        mephiAuthCount = authorsMEPHI.Length;
                        for (int i = 0; i < mephiAuthCount; i++)
                        {
                            authorsMEPHIArray[i, 0] = authorsMEPHI.GetValue(i).ToString();
                            authorsMEPHIArray[i, 1] = tempAffil.Replace("*", "");
                        }
                    }
                    else
                    {
                        int eFlag = 0;
                        foreach (string s in wosArray)
                        {
                            if (eFlag == 0)
                            {
                                Console.WriteLine("# WoS:{0}|", s);
                                string regexGen = "";
                                //searchStringsSplitted
                                string[] searchSS = s.Split(' ');
                                for (int i = 0; i < searchSS.Count(); i++)
                                {
                                    regexGen = regexGen + "" + searchSS[i];
                                }
                                string regexGenRep = "(?i)" + regexGen.Replace("*", "(.*?)");

                                foreach (string affil in affilWosSplit)
                                {
                                    Match authorsWoS = Regex.Match(affil, affilAuthorsPattern, RegexOptions.Singleline);
                                    Match affilMatch = Regex.Match(affil, regexGenRep, RegexOptions.Singleline);

                                    if (authorsWoS.Length == 0)
                                    {
                                        //string with affiliation isn't normal
                                    }
                                    else
                                    {
                                        if (affilMatch.Length > 1)
                                        {
                                            authorsMEPHI = authorsWoS.Groups[1].Value.Replace(".", "").Replace(" ", "").Split(';');

                                            //add each author and his afiliation to array
                                            for (int i = 0; i < authorsMEPHI.Length; i++)
                                            {
                                                authorsMEPHIArray[i, 0] = authorsMEPHI.GetValue(i).ToString();
                                                authorsMEPHIArray[i, 1] = authorsWoS.Groups[2].Value.Replace("*", "").TrimStart();
                                            }
                                            eFlag = 1;
                                        }
                                    }
                                }
                            }
                        }
                        mephiAuthCount = authorsMEPHI.Length;
                    }
                    //checking all authors & mephi authors with magic
                    allAuthorsStr = allAuthors.Groups[1].Value
                        .Replace('\r', ' ')
                        .Replace('\n', ' ')
                        .Replace(".", "")
                        .Replace(" ", "")
                        .Split(';');

                    string[] authorsExceptMEPHI = allAuthorsStr.Except(authorsMEPHI.ToArray()).ToArray();

                    if (authorsExceptMEPHI.Count() >= 4)
                    {
                        for (int i = 0; i <= 3; i++) { authorsNOTMEPHI.Add(authorsExceptMEPHI.GetValue(i).ToString()); };
                    }
                    else if (authorsExceptMEPHI.Count() == 0) { }
                    else
                    {
                        for (int i = 0; i < authorsExceptMEPHI.Count(); i++)
                        {
                            authorsNOTMEPHI.Add(authorsExceptMEPHI.GetValue(i).ToString());
                        }
                    }
                    string[] authorsNOTMEPHIarray = authorsNOTMEPHI.ToArray();

                    //crutch for a title
                    string newTitle = "";
                    string[] titles_temp = title.Groups[1].Value.Split('\n');
                    foreach (string title_temp in titles_temp)
                    {
                        newTitle = newTitle + title_temp;
                    }

                    Console.WriteLine("------------Current record info------------");
                    Console.WriteLine("#Title: {0}", newTitle.Replace("   ", " "));
                    Console.WriteLine("#MEPHI authors number: {0}", mephiAuthCount);
                    Console.WriteLine("#Times Cited: {0}", cited.Groups[1].Value);
                    Console.WriteLine("#ISSN: {0}", issn.Groups[1].Value);
                    Console.WriteLine("#URL: {0}", url.Groups[1].Value);
                    Console.WriteLine("#Journal: {0}", journal.Groups[1].Value);
                    Console.WriteLine("#Volume: {0}", volume.Groups[2].Value);
                    Console.WriteLine("#Issue: {0}", issue.Groups[2].Value);
                    Console.WriteLine("#DOI: {0}", doi.Value);
                    Console.WriteLine("#Pages: {0}", pages.Value);
                    Console.WriteLine("#Year: {0}", year.Value);
                    Console.WriteLine("#Collaboration: {0}", collab.Groups[1].Value);
                    Console.WriteLine("#Document Type: {0}", doctype.Groups[1].Value);

                    IrbisWork(
                        title.Groups[1].Value,
                        authorsNOTMEPHIarray,
                        authorsMEPHIArray,
                        mephiAuthCount,
                        cited.Groups[1].Value,
                        issn.Groups[1].Value,
                        url.Groups[1].Value,
                        journal.Groups[1].Value,
                        volume.Groups[2].Value,
                        issue.Groups[2].Value,
                        doi.Value,
                        pages.Value,
                        year.Value,
                        abstr.Groups[1].Value.Replace('\r', ' ').Replace('\n', ' '),
                        collab.Groups[1].Value,
                        cWosScopus,
                        doctype.Groups[1].Value,
                        allAuthorsStr);
                }
                fileTXT.Close();
            }
            else if (from.Contains("alert@scopus.com"))
            {
                //Scopus
                int cWosScopus = 1;

                //get info entry by entry
                string numberPattern = @"In the overview below, you can see the [1-9][0-9]* result for this Search Alert:";
                string urlPattern = @"scopus.com/alert/results/record.url(.*?)SingleRecordEmailAlert";
                Match number = Regex.Match(email.Text, numberPattern, RegexOptions.Singleline);
                MatchCollection urlArray = Regex.Matches(email.Text, urlPattern, RegexOptions.Singleline);
                Console.WriteLine("Number of articles: {0}", number.Groups[1].Value);

                foreach (Match url in urlArray)
                {
                    Console.WriteLine("------------Current record info------------");
                    Console.WriteLine("Launching PhantomJS");

                    //setting parameters and initializing driver
                    var service = PhantomJSDriverService.CreateDefaultService();
                    service.AddArgument("--ignore-ssl-errors=yes");
                    service.AddArgument("--ssl-protocol=TLSv1");
                    service.AddArgument("--load-images=false");
                    //service.AddArgument("--web-security=false");

                    var phantomDriver = new PhantomJSDriver(service);

                    //going to url
                    string nUrl = "http://" + url.ToString() + "&path_choice=31989";
                    phantomDriver.Navigate().GoToUrl(nUrl);

                    //TODO: get rid of the Regexp
                    //cannot really get rid of the regex here, because most of the parameters are located inside of the JS
                    String innerHTML = phantomDriver.FindElementByTagName("html").GetAttribute("innerHTML");
                    string preAuthorsPatternScopus = "<p id=\"authorlist\"(.*?)>(.*?)</p>";
                    string authorsPatternScopus = "title=\"Show Author Details\">(.*?)</a>";
                    string citedPatternScopus = @"Cited by [0-9]* ";
                    string issnPatternScopus = @"<strong>ISSN: </strong>(.*?)</span>";
                    string doiPatternScopus = "<span id=\"recordDOI\">(.*?)</span></span>";
                    string titlePatternScopus = "var titleForHub = \"(.*?)\";";
                    string journalPatternScopus = "var sourceTitleForHub = \"(.*?)\";";
                    string volumePatternScopus = "var volumeForHub = \"(.*?)\";";
                    string issuePatternScopus = "var issueForHub = \"(.*?)\";";
                    string pagesPatternScopus = "var pagesForHub = \"(.*?)\";";
                    string yearPatternScopus = "var yearForHub = \"(.*?)\";";
                    string abstrPatternScopus = "Abstract</h2>\r\n<p(.*?)>(.*?)</p>";
                    string docTypePatternScopus = "<strong>Document Type: </strong>(.*?)</span>";

                    //getting <p> block with all authors
                    Match preAuFound = Regex.Match(innerHTML, preAuthorsPatternScopus, RegexOptions.Singleline);

                    //getting all authors from previous variable
                    MatchCollection allAuthorsSc = Regex.Matches(preAuFound.Groups[2].Value, authorsPatternScopus);

                    Match title = Regex.Match(innerHTML, titlePatternScopus);
                    Match cited = Regex.Match(innerHTML, citedPatternScopus);
                    Match issn = Regex.Match(innerHTML, issnPatternScopus);
                    Match journal = Regex.Match(innerHTML, journalPatternScopus);
                    Match volume = Regex.Match(innerHTML, volumePatternScopus);
                    Match issue = Regex.Match(innerHTML, issuePatternScopus);
                    Match doi = Regex.Match(innerHTML, doiPatternScopus);
                    Match pages = Regex.Match(innerHTML, pagesPatternScopus);
                    Match year = Regex.Match(innerHTML, yearPatternScopus);
                    Match abstr = Regex.Match(innerHTML, abstrPatternScopus, RegexOptions.Singleline);
                    Match doctype = Regex.Match(innerHTML, docTypePatternScopus);

                    //convert all authors ->
                    List<string> allAuthorsScArray = new List<string>();
                    foreach (Match author in allAuthorsSc)
                    {
                        allAuthorsScArray.Add(author.Groups[1].Value.Replace(" ", ""));
                    }


                    //=================================================================================
                    // -Loading ini file, getting search queries and using it to find out scopus code
                    //=================================================================================
                    //loading ini file
                    System.IO.StreamReader searchFile = new System.IO.StreamReader("searching.ini");
                    String searchStrings = searchFile.ReadToEnd();
                    searchFile.Close();

                    int mephiAuthCount = 0;
                    string[] authorsMEPHISc = { };
                    string[,] authorsMEPHIArray = new string[100, 100]; // 100 is a max number of (mephi authors + 4 not mephi authors)
                    List<string> authorsArrayBuff = new List<string>();
                    List<string> authorsNOTMEPHISc = new List<string>();

                    //getting scopus search queries as strings
                    string scopusSearchPattern = @"\[Scopus\]\r\n(.*?)\r\n\[";
                    string[] charsToSplit = new string[] { "\r\n" };
                    Match scopusQueries = Regex.Match(searchStrings, scopusSearchPattern, RegexOptions.Singleline);

                    //getting array of scopus search queries
                    string[] scopusSQ = scopusQueries.Groups[1].Value.Split(charsToSplit, StringSplitOptions.None);

                    string affilPattern = @"\r\n<span><sup>(.*?)&nbsp;\r\n</sup>(.*?)</span>";
                    string oneAffilPattern = "\r\n<p id=\"affiliationlist\"(.*?)\r\n<span>(.*?)</span>\r\n";
                    string stripPattern = "[a-z*?]+<";
                    string affilAuthorsPattern = "title=\"Show Author Details\"(.*?)>(.*?)</a>(.*?)<sup>(.*?)(</span|<font|<img|<a)";
                    //garbage^ author^ garbage^    code^
                    //getting all authors
                    MatchCollection auFound = Regex.Matches(innerHTML, affilAuthorsPattern, RegexOptions.Singleline);

                    //getting affiliations, if there is more than one
                    MatchCollection affilFound = Regex.Matches(innerHTML, affilPattern, RegexOptions.Singleline);

                    //getting affiliation, if there is only one
                    MatchCollection oneAffilFound = Regex.Matches(innerHTML, oneAffilPattern, RegexOptions.Singleline);

                    if ((affilFound.Count == 0) && (oneAffilFound.Count != 0))
                    {
                        Console.WriteLine("only one afiliation, all authors are from mephi");
                        authorsMEPHISc = allAuthorsScArray.ToArray();

                        //add each author and his afiliation to array
                        for (int i = 0; i < allAuthorsScArray.ToArray().Length; i++)
                        {
                            authorsMEPHIArray[i, 0] = allAuthorsScArray.ToArray().GetValue(i).ToString();
                            authorsMEPHIArray[i, 1] = oneAffilFound[0].Groups[2].Value;
                        }
                        mephiAuthCount = authorsMEPHISc.Length;
                    }
                    else
                    {
                        int eFlag = 0;
                        foreach (string s in scopusSQ)
                        {
                            if (eFlag != 1)
                            {
                                foreach (Match affil in affilFound)
                                {
                                    string affilClean = affil.Groups[2].Value.Replace("'", "").Replace(".", "").Replace("-", " ").Replace("/", " ");
                                    if (affilClean.IndexOf(s, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                    {
                                        string auCode = affil.Groups[1].Value;
                                        Console.WriteLine("String \"{0}\" found. Scopus code:{1}", s, auCode);
                                        eFlag = 1;

                                        foreach (Match author in auFound)
                                        {
                                            string stripAuthor = author.Groups[4].Value;
                                            MatchCollection foundCodes = Regex.Matches(stripAuthor, stripPattern, RegexOptions.Singleline);

                                            foreach (Match code in foundCodes)
                                            {
                                                if ((code.Value.Replace("<", "")) == auCode)
                                                {
                                                    authorsArrayBuff.Add(author.Groups[2].Value.Replace(" ", ""));
                                                }
                                            }
                                        }
                                        //add each author from an article
                                        for (int i = 0; i < authorsArrayBuff.ToArray().Length; i++)
                                        {
                                            authorsMEPHIArray[i, 0] = authorsArrayBuff.ToArray().GetValue(i).ToString();
                                            authorsMEPHIArray[i, 1] = affil.Groups[2].Value;
                                        }
                                    }
                                }
                            }
                            authorsMEPHISc = authorsArrayBuff.ToArray();
                        }
                        mephiAuthCount = authorsMEPHISc.Length;
                    }

                    Console.WriteLine("------------Current record info------------");
                    Console.WriteLine("#Title: {0}", title.Groups[1].Value);
                    Console.WriteLine("#MEPHI authors number: {0}", mephiAuthCount);
                    Console.WriteLine("#Times cited: {0}", cited.Value.Remove(0, 9).TrimEnd());
                    Console.WriteLine("#ISSN: {0}", issn.Groups[1].Value);
                    Console.WriteLine("#URL: {0}", nUrl);
                    Console.WriteLine("#Journal: {0}", journal.Groups[1].Value);
                    Console.WriteLine("#Volume: {0}", volume.Groups[1].Value);
                    Console.WriteLine("#Issue: {0}", issue.Groups[1].Value);
                    Console.WriteLine("#DOI: {0}", doi.Groups[1].Value);
                    Console.WriteLine("#Pages: {0}", pages.Groups[1].Value);
                    Console.WriteLine("#Year: {0}", year.Groups[1].Value);
                    Console.WriteLine("#Document Type: {0}", doctype.Groups[1].Value);

                    //checking all authors & mephi authors with magic
                    string[] authorsExceptMEPHISc = allAuthorsScArray.Except(authorsMEPHISc).ToArray();

                    if (authorsExceptMEPHISc.Count() >= 4)
                    {
                        for (int i = 0; i <= 3; i++)
                        {
                            authorsNOTMEPHISc.Add(authorsExceptMEPHISc.GetValue(i).ToString());
                        }
                    }
                    else if (authorsExceptMEPHISc.Count() == 0) { }
                    else
                    {
                        for (int i = 0; i < authorsExceptMEPHISc.Count(); i++)
                        {
                            authorsNOTMEPHISc.Add(authorsExceptMEPHISc.GetValue(i).ToString());
                        }
                    }
                    string[] authorsNOTMEPHIScArray = authorsNOTMEPHISc.ToArray();

                    IrbisWork(
                        title.Groups[1].Value,
                        authorsNOTMEPHIScArray,
                        authorsMEPHIArray,
                        mephiAuthCount,
                        cited.Value.Remove(0, 9).TrimEnd(),
                        issn.Groups[1].Value,
                        nUrl,
                        journal.Groups[1].Value,
                        volume.Groups[1].Value,
                        issue.Groups[1].Value,
                        doi.Groups[1].Value,
                        pages.Groups[1].Value,
                        year.Groups[1].Value,
                        abstr.Groups[2].Value,
                        "0",
                        cWosScopus,
                        doctype.Groups[1].Value,
                        allAuthorsScArray.ToArray());

                    Console.WriteLine("Closing Chrome");
                    phantomDriver.Close();
                    phantomDriver.Quit();
                }
                fileTXT.Close();
            }
            else
            {
                Console.WriteLine("This email is not from Scopus or WoS");
            }
        }

        //===================================================
        // Major function for Irbis processing
        //===================================================
        private static void IrbisWork(
            string title,
            string[] authorsNotMephi,
            string[,] authorsMEPHIArray,
            int mephiAuthCount,
            string cited, string issn,
            string url, string journal,
            string volume, string issue,
            string doi, string pages,
            string year, string abstr,
            string collab, int cWosScopus,
            string doctype, string[] allAuthorsArray)
        {
            try
            {
                using (ManagedClient64 client = new ManagedClient64())
                {
                    Console.WriteLine("------------Irbis------------");

                    //connecting
                    var irbisSettings = ConfigurationManager.GetSection("irbisSettings") as NameValueCollection;
                    string nDB = irbisSettings["db"].ToString();
                    string nHost = irbisSettings["host"].ToString();
                    string nPort = irbisSettings["port"].ToString();
                    string nUser = irbisSettings["user"].ToString();
                    string nPass = irbisSettings["pass"].ToString();
                    string nConnect = "host=" + nHost + ";port=" + nPort + ";user=" + nUser + ";password=" + nPass + ";db=" + nDB;
                    client.ParseConnectionString(nConnect);
                    client.Connect();

                    //getting rid of whitespaces
                    doi = doi.Trim();
                    year = year.Trim();

                    //=========================================================
                    // Search block.
                    // Search by MFN first, then use keywords from the title
                    //=========================================================
                    string wordPattern = @"\b[A-Za-z]{4}[A-Za-z]*\b";
                    MatchCollection wordArray = Regex.Matches(title, wordPattern);
                    Console.WriteLine("word count:{0};", wordArray.Count);

                    //forming search string with keywords
                    string searchTitleIrbis = "";
                    foreach (Match word in wordArray)
                    {
                        searchTitleIrbis = searchTitleIrbis + " * (\"K=" + RemoveDiacritics(word.Value) + "$\")";
                    }
                    string searchString = searchTitleIrbis.Remove(0, 3) + " * (F)";

                    int[] foundRecordsbyDOI = client.Search("\"DOI={0}$\"", doi);
                    int[] foundRecords = client.Search("{0}", searchString);
                    int recordsFoundbyDOI = foundRecordsbyDOI.Length;
                    int recordsToShow = foundRecords.Length;
                    Console.WriteLine("records found by DOI:{0}", recordsFoundbyDOI);

                    int[] dublArrayMFN = new int[recordsFoundbyDOI];
                    int[] dublArray = new int[recordsToShow];
                    List<int> MFNproblems = new List<int>();

                    if (recordsFoundbyDOI == 1)
                    {
                        //adding found record's mfn to the array
                        MFNproblems.Add(Convert.ToInt32(foundRecordsbyDOI.GetValue(0)));

                        dublArrayMFN[0] = 0;
                        //One record has been found, its definitely the same article, so we are just adding missing fields in it
                        editIrbisRecord(client, mephiAuthCount, cited, url, doi, collab, cWosScopus, doctype, MFNproblems, dublArrayMFN);
                    }
                    if ((recordsFoundbyDOI > 1) && (doi.Length != 0))
                    {
                        Console.WriteLine("more than one record has been found, send email");
                        tooMany(client, title, url, doi, year, MFNproblems);
                    }

                    if (((recordsFoundbyDOI == 0) && (doi.Length != 0)) || ((recordsFoundbyDOI > 1) && (doi.Length == 0)))
                    {
                        Console.WriteLine("no records found, check by searching with keywords");
                        Console.WriteLine("searchTitleIrbis:{0}|", searchString);
                        Console.WriteLine("records found:{0}", recordsToShow);

                        for (int i = 0; i < recordsToShow; i++)
                        {
                            int currentMFN = foundRecords[i];

                            //reading current record
                            IrbisRecord record = client.ReadRecord(currentMFN);
                            string currentTitle = record.FM("200", 'a');
                            string currentDOI = record.FM("254");
                            string currentYear = record.FM("463", 'j');
                            Console.WriteLine("---");
                            Console.WriteLine("MFN={0}, Title={1}, DOI={2}, Year={3}", currentMFN, currentTitle, currentDOI, currentYear);

                            //checking if record already exists
                            if (doi.Length == 0)
                            {
                                Console.WriteLine("# doi is null, check year");
                                if ((year.Length == 0) || (currentYear.Length == 0))
                                {
                                    dublArray[i] = 0;
                                    MFNproblems.Add(currentMFN);
                                    Console.WriteLine("# year/currentYear is null, cant distinguish if the same");
                                }
                                else
                                {
                                    if (year == currentYear)
                                    {
                                        dublArray[i] = 0;
                                        MFNproblems.Add(currentMFN);
                                        Console.WriteLine("# year is the same, doi is empty, probably the same article");
                                    }
                                    else
                                    {
                                        dublArray[i] = 1;
                                        Console.WriteLine("# year is not the same, doi is empty, not the same article");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("# doi={0}, currentDOI={1}.", doi, currentDOI);
                                if (doi == currentDOI)
                                {
                                    dublArray[i] = 0;
                                    MFNproblems.Add(currentMFN);
                                    Console.WriteLine("# doi is the same, its the same article");
                                }
                                else
                                {
                                    dublArray[i] = 1;
                                    Console.WriteLine("# doi is not the same, not the same article");
                                }
                            }
                        }

                        Console.WriteLine("------------Total------------");
                        //if all the array elements are "1": ==
                        //create a new record
                        if (dublArray.Sum() == recordsToShow)
                        {
                            createIrbisRecord(client, title, authorsNotMephi, authorsMEPHIArray, mephiAuthCount, cited, issn, url, journal, volume, issue, doi, pages, year, abstr, collab, cWosScopus, doctype, allAuthorsArray);
                        }
                        else
                        {
                            //we need to count the number of "0" in dublArray
                            var numArray = dublArray.Where(num => num == 0);
                            if (numArray.Count() == 1)
                            {
                                //only one suitable article found, get it's MFN and add stuff in it
                                editIrbisRecord(client, mephiAuthCount, cited, url, doi, collab, cWosScopus, doctype, MFNproblems, dublArray);
                            }
                            else
                            {
                                //too many suitable articles, send email
                                Console.WriteLine("more than one record has been found, send email");
                                tooMany(client, title, url, doi, year, MFNproblems);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        //===================================================================
        // createIrbisRecord() function adds a record to the IRBIS database
        //===================================================================
        public static void createIrbisRecord(ManagedClient64 client, string title, string[] authorsNotMephi, string[,] authorsMEPHIArray, int mephiAuthCount,
            string cited, string issn, string url, string journal, string volume, string issue, string doi,
            string pages, string year, string abstr, string collab, int cWosScopus, string doctype, string[] allAuthorsArray)
        {
            Console.WriteLine("# Creating a new record");
            //creating a new record
            IrbisRecord newRecord = new IrbisRecord();

            int aSumm = authorsNotMephi.Length + mephiAuthCount;
            if (aSumm >= 4)
            {
                //add all authors to the 701 field
                foreach (string author in authorsNotMephi)
                {
                    newRecord.AddField("701", 'A', author);
                }
                for (int k = 0; k < mephiAuthCount; k++)
                {
                    newRecord.AddField("701", 'A', authorsMEPHIArray[k, 0], 'P', authorsMEPHIArray[k, 1]);
                }
            }
            else //the number of authors is less than 4
            {
                int amCount = mephiAuthCount;
                int anmCount = authorsNotMephi.Length;

                //inserting 4 or less authors NOT from mephi
                //if only 1 author found, place him
                if (anmCount == 1)
                {
                    newRecord.AddField("700", 'A', authorsNotMephi.GetValue(0).ToString());
                    for (int k = 0; k < mephiAuthCount; k++)
                    {
                        newRecord.AddField("701", 'A', authorsMEPHIArray[k, 0], 'P', authorsMEPHIArray[k, 1]);
                    }
                }
                else if (anmCount == 0)//if no NotMEPHI authors found, add only mephi
                {
                    newRecord.AddField("700", 'A', authorsMEPHIArray[0, 0], 'P', authorsMEPHIArray[0, 1]);
                    for (int k = 1; k < mephiAuthCount; k++)
                    {
                        newRecord.AddField("701", 'A', authorsMEPHIArray[k, 0], 'P', authorsMEPHIArray[k, 1]);
                    }
                }
                else if ((anmCount > 1) && (anmCount < 4))//if 0<x<4, insert the first one in 700, and the rest in 701
                {
                    newRecord.AddField("700", 'A', authorsNotMephi.GetValue(0).ToString());
                    for (int i = 1; i < anmCount; i++)
                    {
                        newRecord.AddField("701", 'A', authorsNotMephi.GetValue(i).ToString());
                    }
                    for (int k = 0; k < mephiAuthCount; k++)
                    {
                        newRecord.AddField("701", 'A', authorsMEPHIArray[k, 0], 'P', authorsMEPHIArray[k, 1]);
                    }
                }
            }

            //===================================================================
            //add first 4 or less authors to the debug fields (721 and 722)
            if (allAuthorsArray.Length >= 4)
            {
                for (int k = 0; k < 4; k++)
                {
                    newRecord.AddField("722", allAuthorsArray.GetValue(k).ToString());
                }
            }
            else
            {
                newRecord.AddField("721", allAuthorsArray.GetValue(0).ToString());
                for (int k = 1; k < allAuthorsArray.Length; k++)
                {
                    newRecord.AddField("722", allAuthorsArray.GetValue(k).ToString());
                }
            }
            //===================================================================

            newRecord.AddField("200", 'A', RemoveDiacritics(title));
            newRecord.AddField("463", 'C', journal, 'J', year, 'V', volume, 'S', pages, 'Q', issue);
            newRecord.AddField("920", "ASP");
            newRecord.AddField("900", 'B', "08");
            newRecord.AddField("254", doi);
            newRecord.AddField("11", issn);
            newRecord.AddField("951", 'I', url);
            newRecord.AddField("331", abstr);
            newRecord.AddField("259", collab);

            if (cWosScopus == 0) //wos
            {
                newRecord.AddField("255", 'A', cited);
                newRecord.AddField("256", 'A', "1", 'C', string.Format(DateTime.Now.ToString("yyyyMMdd")));
                newRecord.AddField("257", 'A', doctype);
                newRecord.AddField("705", 'A', mephiAuthCount.ToString());
            }
            else if (cWosScopus == 1) //scopus
            {
                newRecord.AddField("255", 'B', cited);
                newRecord.AddField("256", 'B', "1", 'D', string.Format(DateTime.Now.ToString("yyyyMMdd")));
                newRecord.AddField("257", 'B', doctype);
                newRecord.AddField("705", 'B', mephiAuthCount.ToString());
            }

            //sending this record to the server
            client.WriteRecord(newRecord, false, true);
            Console.WriteLine("# New record has been created");

            //adding info to the email template (case of a record creating)
            articlesHTML.Add("<tr><td>" + mephiAuthCount + "</td><td><a href=\"" + url + "\">"
                                        + title + "</a></td><td>"
                                        + journal + "</td><td>"
                                        + year + "</td><td>"
                                        + doi + "</td></tr>");

            //clearing arrays
            Array.Clear(authorsMEPHIArray, 0, authorsMEPHIArray.Length);
            Array.Clear(authorsNotMephi, 0, authorsNotMephi.Length);
        }

        //===================================================================
        // editIrbisRecord() function changes a record in the IRBIS database
        //===================================================================
        public static void editIrbisRecord(ManagedClient64 client, int mephiAuthCount, string cited, string url, string doi,
            string collab, int cWosScopus, string doctype, List<int> MFNproblems, int[] dublArray)
        {
            //only one suitable article found, get it's MFN and add stuff in it
            Console.WriteLine("one suitable article has been found, get it's MFN and add stuff in it");

            //searching for zero in dubl array, and finding MFN that corresponds to that index
            int MFNn = Array.IndexOf(dublArray, 0);
            int mfn = (int)MFNproblems.ToArray().GetValue(MFNn);
            Console.WriteLine("mfn:{0}", mfn.ToString());

            IrbisRecord addToRecord = client.ReadRecord(mfn);
            string currentDOI = addToRecord.FM("254"); //getting record's doi

            //adding publication date, db existence mark, number of times it was cited, doc.type, collaborations and number of mephi authors
            if (cWosScopus == 0) // wos
            {
                addToRecord.SetSubField("255", 'A', cited).SetSubField("256", 'A', "1").SetSubField("257", 'A', doctype)
                           .SetSubField("256", 'C', string.Format(DateTime.Now.ToString("yyyyMMdd")))
                           .SetSubField("705", 'A', mephiAuthCount.ToString())
                           .SetField("259", collab);
                if ((doi.Length != 0) && (currentDOI.Length == 0))
                {
                    addToRecord.SetField("254", doi);
                }
            }
            else //scopus
            {
                addToRecord.SetSubField("255", 'B', cited).SetSubField("256", 'B', "1").SetSubField("257", 'B', doctype)
                           .SetSubField("256", 'D', string.Format(DateTime.Now.ToString("yyyyMMdd")))
                           .SetSubField("705", 'B', mephiAuthCount.ToString());
                if ((doi.Length != 0) && (currentDOI.Length == 0))
                {
                    addToRecord.SetField("254", doi);
                }
                if (url.Length != 0)
                {
                    addToRecord.SetSubField("951", 'I', url);
                }
            }
            client.WriteRecord(addToRecord, false, true);
            Console.WriteLine("# Done");

            IrbisRecord record = client.ReadRecord(mfn);
            string currTitle = record.FM("200", 'a');
            string currDOI = record.FM("254");
            string currYear = record.FM("463", 'j');
            articlesCorrected.Add("<tr><td>" + mfn + "</td><td>"
                                             + currTitle + "</td><td>"
                                             + currDOI + "</td><td>"
                                             + currYear + "</td></tr>");
        }

        //=============================================
        // tooMany() function forms an array for email
        //=============================================
        public static void tooMany(ManagedClient64 client, string title, string url, string doi, string year, List<int> MFNproblems)
        {
            //too many suitable articles, form an array for email
            articlesHTMLproblems.Add("<tr><td class=\"orig\">new</td><td class=\"orig\"><a href=\""
                                                + url + "\">" + title + "</a></td><td class=\"orig\">"
                                                + doi + "</td><td class=\"orig\">"
                                                + year + "</td></tr>");
            foreach (int mfn in MFNproblems)
            {
                IrbisRecord record = client.ReadRecord(mfn);
                string currTitle = record.FM("200", 'a');
                string currDOI = record.FM("254");
                string currYear = record.FM("463", 'j');
                articlesHTMLproblems.Add("<tr><td>" + mfn + "</td><td>"
                                                    + currTitle + "</td><td>"
                                                    + currDOI + "</td><td>"
                                                    + currYear + "</td></tr>");
            }
        }

        //=========================================================================================
        // Functions for diacritics removing, mail processing and logging
        //=========================================================================================
        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();
            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }
            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        private static string JoinAddresses(IList<MailBox> mailboxes)
        {
            return string.Join(",", new List<MailBox>(mailboxes).ConvertAll(m => string.Format("{0} <{1}>", m.Name, m.Address)).ToArray());
        }

        private static string JoinAddresses(IList<MailAddress> addresses)
        {
            StringBuilder builder = new StringBuilder();

            foreach (MailAddress address in addresses)
            {
                if (address is MailGroup)
                {
                    MailGroup group = (MailGroup)address;
                    builder.AppendFormat("{0}: {1};, ", group.Name, JoinAddresses(group.Addresses));
                }
                if (address is MailBox)
                {
                    MailBox mailbox = (MailBox)address;
                    builder.AppendFormat("{0} <{1}>, ", mailbox.Name, mailbox.Address);
                }
            }
            return builder.ToString();
        }
    }

    //class for mirror logging
    class ConsoleCopy : IDisposable
    {
        FileStream fileStream;
        StreamWriter fileWriter;
        TextWriter doubleWriter;
        TextWriter oldOut;

        class DoubleWriter : TextWriter
        {
            TextWriter one;
            TextWriter two;

            public DoubleWriter(TextWriter one, TextWriter two)
            {
                this.one = one;
                this.two = two;
            }

            public override Encoding Encoding
            {
                get { return one.Encoding; }
            }

            public override void Flush()
            {
                one.Flush();
                two.Flush();
            }

            public override void Write(char value)
            {
                one.Write(value);
                two.Write(value);
            }
        }

        public ConsoleCopy(string path)
        {
            oldOut = Console.Out;

            try
            {
                fileStream = File.Create(path);
                fileWriter = new StreamWriter(fileStream);
                fileWriter.AutoFlush = true;
                doubleWriter = new DoubleWriter(fileWriter, oldOut);
            }
            catch (Exception e)
            {
                Console.WriteLine("Cannot open the file");
                Console.WriteLine(e.Message);
                return;
            }
            Console.SetOut(doubleWriter);
        }

        public void Dispose()
        {
            Console.SetOut(oldOut);
            if (fileWriter != null)
            {
                fileWriter.Flush();
                fileWriter.Close();
                fileWriter = null;
            }
            if (fileStream != null)
            {
                fileStream.Close();
                fileStream = null;
            }
        }
    }
}