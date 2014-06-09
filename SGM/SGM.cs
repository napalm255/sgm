using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Collections;
using System.Text;
using System.Net;
using System.Diagnostics;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.Net.Mail;
using System.ComponentModel;
using System.Configuration;

namespace Mgmt.SGM
{
    public class SGMService: IDisposable
    {
        // Pointer to an external unmanaged resource.
        private IntPtr handle;
        // Other managed resource this class uses.
        private Component component = new Component();
        // Track whether Dispose has been called.
        private bool disposed = false;

        Dictionary<string, string> iSettings = new Dictionary<string, string>();
        Dictionary<string, string> iExclude = new Dictionary<string, string>();

        public SGMService()
        {
            return;
        }

        public void Start()
        {
            try
            {
                // Load Configuration
                LoadConfig();
                // Process Groups
                Query();

                // Convert Seconds to Milliseconds
                //int ThreadSleep = Convert.ToInt16(iSettings["interval"]) * 1000;
                // Put the thread to sleep
                //Thread.Sleep(ThreadSleep);
            }
            catch { /* DEBUG */ }
        }

        public void Stop()
        {
            return;
        }

        public void Query()
        {
            try
            {
                // Open SQL Connection
                using (SqlConnection dbConn = SqlConnect())
                {
                    // Open Database Connection
                    dbConn.Open();

                    // Process Groups
                    try
                    {
                        SqlCommand sCmd = new SqlCommand("SELECT * FROM sgmMonitor WHERE sInclude = 1 AND sType = 'group'", dbConn);
                        using (SqlDataReader dbReader = sCmd.ExecuteReader())
                        {
                            while (dbReader.Read())
                            {
                                if (dbReader["sType"].ToString() == "group")
                                {
                                    CheckGroup(dbReader["sObject"].ToString());
                                }
                            }
                            dbReader.Close();
                            dbReader.Dispose();
                        }
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch { /* DEBUG */ }

                    // Process OUs
                    try
                    {
                        SqlCommand sCmd = new SqlCommand("SELECT * FROM sgmMonitor WHERE sInclude = 1 AND sType = 'ou'", dbConn);
                        using (SqlDataReader dbReader = sCmd.ExecuteReader())
                        {
                            while (dbReader.Read())
                            {
                                try
                                {
                                    DirectoryEntry objADroot = new DirectoryEntry(@"LDAP://rootDSE");
                                    string pathRoot = objADroot.Invoke("GET", "defaultNamingContext").ToString();
                                    string pathOU = "LDAP://" + dbReader["sObject"] + "," + pathRoot;

                                    DirectoryEntry objADAM = new DirectoryEntry(pathOU);
                                    objADAM.RefreshCache();

                                    DirectorySearcher objSearchADAM = new DirectorySearcher(objADAM);
                                    objSearchADAM.Filter = "(&(objectClass=group))";
                                    objSearchADAM.SearchScope = SearchScope.Subtree;

                                    SearchResultCollection objSearchResults = objSearchADAM.FindAll();

                                    if (objSearchResults.Count != 0)
                                    {
                                        foreach (SearchResult objResult in objSearchResults)
                                        {
                                            DirectoryEntry objGroupEntry = objResult.GetDirectoryEntry();
                                            if (!iExclude.ContainsKey(objGroupEntry.Properties["distinguishedName"].Value.ToString().Replace("," + pathRoot, "").ToLower()))
                                            {
                                                CheckGroup(objGroupEntry.Properties["distinguishedName"].Value.ToString());
                                            }
                                            objGroupEntry.Close();
                                            objGroupEntry.Dispose();
                                            objGroupEntry = null;
                                        }
                                    }

                                    objSearchResults.Dispose();
                                    objSearchResults = null;
                                    objSearchADAM.Dispose();
                                    objADAM.Close();
                                    objADAM.Dispose();
                                    objADAM = null;
                                    pathOU = null;
                                    pathRoot = null;
                                    objADroot.Close();
                                    objADroot.Dispose();
                                    objADroot = null;
                                }
                                catch { /* DEBUG */ }
                            }
                            dbReader.Close();
                            dbReader.Dispose();
                        }
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch
                    {
                        /* DEBUG */
                    }

                    // Process SQL Roles
                    try
                    {
                        SqlCommand sCmd = new SqlCommand("SELECT * FROM sgmMonitor WHERE sInclude = 1 AND sType = 'sql:roles'", dbConn);
                        using (SqlDataReader dbReader = sCmd.ExecuteReader())
                        {
                            while (dbReader.Read())
                            {
                                try
                                {
                                    ArrayList sqlRoles = getSqlRoles(dbReader["sObject"].ToString());
                                    foreach (string role in sqlRoles)
                                    {
                                        CheckSqlRole(dbReader["sObject"].ToString(), role);
                                    }
                                    sqlRoles.Clear();
                                    sqlRoles = null;
                                }
                                catch { /* DEBUG */ }
                            }
                            dbReader.Close();
                            dbReader.Dispose();
                        }
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch { /* DEBUG */ }
                }
            }
            catch { /* DEBUG */ }
        }

        private ArrayList getSqlRoles(string oConStr)
        {
            // Define Variables
            ArrayList aryRoles = new ArrayList();

            // Open SQL Connection
            using (SqlConnection dbConn = SqlConnect(oConStr))
            {
                // Open Database Connection
                dbConn.Open();

                // Process Groups
                try
                {
                    SqlCommand sCmd = new SqlCommand("exec sp_helpsrvrole", dbConn);
                    using (SqlDataReader dbReader = sCmd.ExecuteReader())
                    {
                        while (dbReader.Read())
                        {
                            aryRoles.Add(dbReader[0].ToString());
                        }
                        dbReader.Close();
                        dbReader.Dispose();
                    }
                    sCmd.Dispose();
                    sCmd = null;
                }
                catch
                {
                    /* DEBUG */
                    return null;
                }

                // Close Database Connection
                dbConn.Close();
                dbConn.Dispose();
            }
            return aryRoles;
        }

        private string[] getSqlRoleMembers(string oConStr, string oRole)
        {
            // Define Variables
            string[] theMembers = new string[3];

            // Open SQL Connection
            using (SqlConnection dbConn = SqlConnect(oConStr))
            {
                // Open Database Connection
                dbConn.Open();
                
                // Process Groups
                try
                {
                    string theServer = getSrvNameFrmConStr(oConStr);

                    theMembers[0] = oRole + "@" + theServer;
                    theMembers[1] = oRole + "@" + theServer;

                    SqlCommand sCmd = new SqlCommand("exec sp_helpsrvrolemember '" + oRole + "'", dbConn);
                    using (SqlDataReader dbReader = sCmd.ExecuteReader())
                    {
                        while (dbReader.Read())
                        {
                            theMembers[2] += "+" + dbReader[1].ToString() + ";";
                        }
                        dbReader.Close();
                        dbReader.Dispose();
                    }
                    sCmd.Dispose();
                    sCmd = null;
                    theServer = null;
                }
                catch { /* DEBUG */ }

                // Close Database Connection
                dbConn.Close();
                dbConn.Dispose();
            }
            return theMembers;
        }

        private string getSrvNameFrmConStr(string oConStr)
        {
            string theServer = null;
            string[] aryProvider = oConStr.Split(new char[1] { ';' });
            foreach (string item in aryProvider)
            {
                if (item.ToLower().Contains("server="))
                {
                    theServer = item.ToLower().Replace("server=", "");
                }
            }
            aryProvider = null;

            return theServer;
        }

        private void CheckSqlRole(string oConStr, string oRole)
        {
            // Define Variables
            string nMembers = "";
            string rMembers = "";
            string oGroup = oRole + "@" + getSrvNameFrmConStr(oConStr);

            try
            {
                // Log
                //EventLog.WriteEntry("SGM", "Checking Group...", EventLogEntryType.Information);

                // Look up group
                try
                {
                    string[] qGroup = getGroupEntry(oGroup);
                    if (qGroup == null)
                    {
                        string[] gMem = getSqlRoleMembers(oConStr, oRole);
                        UpdateDatabase(gMem[0], gMem[2]);
                        gMem = null;
                    }
                    else
                    {
                        rMembers = qGroup[1];
                        string[] gMem = getSqlRoleMembers(oConStr, oRole);
                        string[] groupMembers = gMem[2].Split(';');
                        foreach (string member in groupMembers)
                        {
                            if (rMembers.Contains(member))
                            {
                                rMembers = rMembers.Replace(member + ";", "");
                            }
                            else
                            {
                                nMembers += member + ";";
                            }
                        }
                        if (nMembers != "" || rMembers != "")
                        {
                            UpdateDatabase(gMem[0], gMem[2], nMembers, rMembers);
                            nMembers = nMembers.Replace("&#39;", "'");
                            rMembers = rMembers.Replace("&#39;", "'");
                            gMem[2] = gMem[2].Replace("&#39;", "'");
                            SendMail(gMem[1], gMem[0], nMembers.Replace(";", "<br>"), rMembers.Replace(";", "<br>"), gMem[2].Replace(";", "<br>"));
                        }

                        groupMembers = null;
                        gMem = null;
                    }
                    qGroup = null;
                }
                catch { /* DEBUG */ }
            }
            catch { /* DEBUG */ }

            nMembers = null;
            rMembers = null;
            oGroup = null;
        }

        private void CheckGroup(string oGroup)
        {
            // Define Variables
            string nMembers = "";
            string rMembers = "";

            try
            {
                // Look up group
                try
                {
                    string[] qGroup = getGroupEntry(oGroup);
                    if (qGroup == null)
                    {
                        string[] gMem = getGroupMembers(oGroup);
                        UpdateDatabase(gMem[0], gMem[2]);
                        gMem = null;
                    }
                    else
                    {
                        rMembers = qGroup[1];
                        string[] gMem = getGroupMembers(oGroup);
                        string[] groupMembers = gMem[2].Split(';');
                        foreach (string member in groupMembers)
                        {
                            if ( (member != "") && (rMembers.Contains("+" + member + ";")) )
                            {
                                rMembers = rMembers.Replace("+" + member + ";", "");
                            }
                            else if (member != "")
                            {
                                nMembers += member + ";";
                            }
                        }

                        if (nMembers != "" || rMembers != "")
                        {
                            UpdateDatabase(gMem[0], gMem[2], nMembers, rMembers);
                            nMembers = nMembers.Replace("&#39;", "'");
                            rMembers = rMembers.Replace("&#39;", "'");
                            gMem[2] = gMem[2].Replace("&#39;", "'");
                            SendMail(gMem[1], gMem[0], nMembers.Replace(";", "<br>").Replace("+", ""), rMembers.Replace(";", "<br>").Replace("+", ""), gMem[2].Replace(";", "<br>").Replace("+", ""));
                        }
                        nMembers = null;
                        rMembers = null;
                        groupMembers = null;
                        gMem = null;
                    }
                    qGroup = null;
                }
                catch { /* DEBUG */ }
            }
            catch { /* DEBUG */ }

            nMembers = null;
            rMembers = null;
        }

        private string[] getGroupMembers(string oGroup)
        {
            // Define Variables
            string[] theMembers = new string[3];

            try
            {
                // Lookup
                DirectoryEntry objADroot = new DirectoryEntry(@"LDAP://rootDSE");
                string pathRoot = objADroot.Invoke("GET", "defaultNamingContext").ToString();
                string pathOU;
                if (oGroup.Contains(pathRoot))
                {
                    pathOU = "LDAP://" + oGroup.Replace("," + pathRoot, ",") + pathRoot;
                }
                else
                {
                    pathOU = "LDAP://" + oGroup + "," + pathRoot;
                }

                DirectoryEntry objAD = new DirectoryEntry(pathOU);
                objAD.RefreshCache();

                DirectorySearcher objSrch = new DirectorySearcher(objAD);
                objSrch.Filter = "(&(objectClass=*))";

                objSrch.SearchScope = SearchScope.Subtree;
                SearchResultCollection objSearchResults  = objSrch.FindAll();

                foreach (SearchResult rs in objSearchResults)
                {
                    theMembers[0] = rs.Path.Substring(7);
                    string[] aryPath = rs.Path.Substring(7).Split(',');
                    theMembers[1] = aryPath[0].Substring(3);

                    ResultPropertyCollection resultPropColl = rs.Properties;
                    foreach (Object memberColl in resultPropColl["member"])
                    {
                        DirectoryEntry gpMemberEntry = new DirectoryEntry("LDAP://" + memberColl);
                        try
                        {
                            if (gpMemberEntry.Name.ToLower().Contains("ou="))
                            {
                                theMembers[2] += gpMemberEntry.Name.Substring(3) + " (OU);";
                            }
                            else
                            {
                                theMembers[2] += gpMemberEntry.Properties["samAccountName"].Value.ToString().Replace("'","&#39;") + ";";
                            }
                        }
                        catch
                        {
                            /* DEBUG */
                            return null;
                        }
                        gpMemberEntry.Close();
                        gpMemberEntry.Dispose();
                        gpMemberEntry = null;
                    }
                    resultPropColl = null;
                    aryPath = null;
                }

                // Clean up variables
                objSearchResults.Dispose();
                objSearchResults = null;
                objSrch.Dispose();
                objSrch = null;
                objAD.Close();
                objAD.Dispose();
                objAD = null;
                pathRoot = null;
                pathOU = null;
                objADroot.Close();
                objADroot.Dispose();
                objADroot = null;
            }
            catch
            {
                /* DEBUG */
                return null;
            }

            return theMembers;
        }

        private string[] getGroupEntry(string oGroup)
        {
            // Define Variables
            string[] gEntry = new string[4];

            try
            {
                // Open SQL Connection
                using (SqlConnection dbConn = SqlConnect())
                {
                    // Open Database Connection
                    dbConn.Open();

                    // Process Groups
                    try
                    {
                        SqlCommand sCmd = new SqlCommand("SELECT TOP 1 tID, tGroupname, tMembers, tNewMembers, tRemMembers, tTimestamp FROM sgmTracking WHERE tGroupname LIKE '%" + oGroup + "%' ORDER BY tTimestamp DESC", dbConn);
                        using (SqlDataReader geReader = sCmd.ExecuteReader())
                        {
                            geReader.Read();
                            
                            gEntry.SetValue(geReader["tGroupname"].ToString(), 0);
                            gEntry.SetValue(geReader["tMembers"].ToString(), 1);
                            gEntry.SetValue(geReader["tNewMembers"].ToString(), 2);
                            gEntry.SetValue(geReader["tRemMembers"].ToString(), 3);

                            geReader.Close();
                            geReader.Dispose();
                        }
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch
                    {
                        /* DEBUG */
                        return null;
                    }

                    // Close Database Connection
                    dbConn.Close();
                    dbConn.Dispose();
                }
            }
            catch
            {
                /* DEBUG */
                return null;
            }
            return gEntry;
        }



        private void LoadConfig()
        {
            try
            {
                // Load database connection string from App.config
                iSettings.Add("sqlprovider",  ConfigurationManager.AppSettings["databaseConnectionString"].ToString());

                // Open SQL Connection
                using (SqlConnection dbConn = SqlConnect())
                {
                    // Open Database Connection
                    dbConn.Open();

                    // Load Settings
                    try
                    {
                        SqlCommand sCmd = new SqlCommand("SELECT * FROM sgmSettings", dbConn);
                        using (SqlDataReader dbReader = sCmd.ExecuteReader())
                        {
                            while (dbReader.Read())
                            {
                                try
                                {
                                    iSettings.Add(dbReader["sAttribute"].ToString().ToLower(), dbReader["sValue"].ToString());
                                }
                                catch
                                {
                                    iSettings[dbReader["sAttribute"].ToString().ToLower()] = dbReader["sValue"].ToString();
                                }
                            }
                            dbReader.Close();
                            dbReader.Dispose();
                        }
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch { /* DEBUG */ }


                    // Load Excluded Objects
                    try
                    {
                        // Clear Includes
                        iExclude.Clear();

                        SqlCommand sCmd = new SqlCommand("SELECT * FROM sgmMonitor WHERE sInclude = 0", dbConn);
                        using (SqlDataReader dbReader = sCmd.ExecuteReader())
                        {
                            while (dbReader.Read())
                            {
                                try
                                {
                                    iExclude.Add(dbReader["sObject"].ToString().ToLower(), dbReader["sType"].ToString());
                                }
                                catch
                                {
                                    iExclude[dbReader["sObject"].ToString().ToLower()] = dbReader["sType"].ToString();
                                }
                            }
                            dbReader.Close();
                            dbReader.Dispose();
                        }
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch { /* DEBUG */ }

                    // Close Database Connection
                    dbConn.Close();
                    dbConn.Dispose();
                }
            }
            catch { /* DEBUG */ }
        }

        private ArrayList LoadAlertStaff()
        {
            // Define Variables
            ArrayList iAlertStaff = new ArrayList();

            // Load Alert Staff from Database
            try
            {
                using (SqlConnection dbConn = SqlConnect())
                {
                    // Open Database Connection
                    dbConn.Open();

                    // Load Settings
                    SqlCommand sCmd = new SqlCommand("SELECT * FROM sgmAlertStaff WHERE aEnabled = 1", dbConn);
                    using (SqlDataReader dbReader = sCmd.ExecuteReader())
                    {
                        while (dbReader.Read())
                        {
                            iAlertStaff.Add(dbReader["aEmail"].ToString());
                        }
                        dbReader.Close();
                        dbReader.Dispose();
                    }
                    sCmd.Dispose();
                    sCmd = null;

                    // Close Database Connection
                    dbConn.Close();
                }
            }
            catch { /* DEBUG */ }
            
            // Load Alert Staff from AlertGroup
            try
            {
                StringCollection groupMembers = this.GetEmailAddresses(iSettings["alertgroup"]);
                foreach (string strMember in groupMembers)
                {
                    iAlertStaff.Add(strMember);
                }
                groupMembers = null;
            }
            catch { /* DEBUG */ }

            return iAlertStaff;
        }

        private StringCollection GetEmailAddresses(string strGroup)
        {
            StringCollection emails = new StringCollection();
            try
            {
                DirectorySearcher srch = new DirectorySearcher("(CN=" + strGroup + ")");
                SearchResultCollection coll = srch.FindAll();
                foreach (SearchResult rs in coll)
                {
                    ResultPropertyCollection resultPropColl = rs.Properties;
                    foreach (Object memberColl in resultPropColl["member"])
                    {
                        DirectoryEntry gpMemberEntry = new DirectoryEntry("LDAP://" + memberColl);
                        try
                        {
                            // Loop through all the assigned e-mail addresses
                            foreach (string proxyAddr in gpMemberEntry.Properties["proxyAddresses"])
                            {
                                // Record the primary e-mail address
                                // The primary is prefixed with all uppercase 'SMTP'.
                                if (proxyAddr.StartsWith("SMTP:"))
                                {
                                    emails.Add(proxyAddr.Substring(5).ToLower());
                                }
                            }
                        }
                        catch
                        {
                            /* DEBUG */
                            emails = null;
                        }
                        gpMemberEntry.Close();
                        gpMemberEntry.Dispose();
                        gpMemberEntry = null;
                    }
                    resultPropColl = null;
                }
                coll.Dispose();
                coll = null;
                srch.Dispose();
                srch = null;
            }
            catch { /* DEBUG */ }
            
            return emails;
        }

        private void SendMail(string strGroupName, string strDN, string strNewMembers, string strRemovedMembers, string strCurrentMembers)
        {
            // Log
            EventLog.WriteEntry("SGM", "Sending Email Notification (" + strGroupName + ")...", EventLogEntryType.Information);

            // Process Mail
            using (MailMessage mailMsg = new MailMessage())
            {
                // Load Alert Staff
                ArrayList iTo = LoadAlertStaff();

                // Configure Mail Message
                try
                {
                    mailMsg.IsBodyHtml = true;
                    mailMsg.From = new MailAddress(iSettings["emailfrom"]);
                    foreach (string userEmail in iTo)
                    {
                        if (userEmail != "" || userEmail != null) { mailMsg.To.Add(userEmail); }
                    }
                    mailMsg.Subject = iSettings["emailsubject"];
                    string msgContents = string.Format("<center><table width=\"50%\" cellpadding=\"3\" cellspacing=\"3\"><tr><td style=\"font-family:verdana;font-weight:bold;font-size:12px;text-align:center;background-color:#99C3ED;border-bottom:1px solid #000000;\" colspan=\"2\">{0}</td></tr><tr><td colspan=\"2\" style=\"font-family:verdana;font-size:10px;text-align:center;border-bottom:1px solid #000000;\">{1}</td></tr><tr><td width=\"30%\" align=\"right\" valign=\"top\" style=\"padding-right:10px;font-family:verdana;font-weight:bold;font-size:10px;\">New Members:</td><td valign=\"top\" style=\"font-family:verdana;font-weight:normal;font-size:10px;\">{2}</td></tr><tr><td width=\"30%\" align=\"right\" valign=\"top\" style=\"padding-right:10px;font-family:verdana;font-weight:bold;font-size:10px;\">Removed Members:</td><td valign=\"top\" style=\"font-family:verdana;font-weight:normal;font-size:10px;\">{3}</td></tr><tr><td width=\"30%\" align=\"right\" valign=\"top\" style=\"padding-right:10px;font-family:verdana;font-weight:bold;font-size:10px;\">Current Members:</td><td valign=\"top\" style=\"font-family:verdana;font-weight:normal;font-size:10px;\">{4}</td></tr></table></center>", strGroupName, strDN, strNewMembers, strRemovedMembers, strCurrentMembers);
                    mailMsg.Body = msgContents;
                    msgContents = null;
                }
                catch { /* DEBUG */ }

                // Send Mail Message
                try
                {
                    SmtpClient mailSmtp = new SmtpClient();
                    mailSmtp.Host = iSettings["emailserver"];
                    //mailSmtp.Credentials = new NetworkCredential("<user>", "<pass>");
                    mailSmtp.Send(mailMsg);
                    mailSmtp = null;
                }
                catch { /* DEBUG */ }
                
                mailMsg.Dispose();
                iTo = null;
            }
        }

        public Boolean SqlAlive()
        {
            try
            {
                using (SqlConnection dbConn = SqlConnect())
                {
                    dbConn.Open();
                    dbConn.Close();
                    dbConn.Dispose();
                    return true;
                }
            }
            catch
            {
                /* DEBUG */
                return false;
            }
        }

        private SqlConnection SqlConnect()
        {
            try
            {
                string theServer = null;
                using (System.Net.NetworkInformation.Ping sqlPing = new System.Net.NetworkInformation.Ping())
                {
                    string[] aryProvider = iSettings["sqlprovider"].Split(new char[1] { ';' });
                    foreach (string item in aryProvider)
                    {
                        if (item.ToLower().Contains("server="))
                        {
                            theServer = item.ToLower().Replace("server=", "");
                        }
                    }
                    aryProvider = null;
                    try
                    {
                        System.Net.NetworkInformation.PingReply sqlReply = sqlPing.Send(theServer);
                        return new SqlConnection(iSettings["sqlprovider"]);
                    }
                    catch
                    {
                        /* DEBUG */
                        return null;
                    }
                }
            }
            catch
            {
                /* DEBUG */
                return null;
            }
        }

        private SqlConnection SqlConnect(string oConnectionString)
        {
            try
            {
                string theServer = null;
                using (System.Net.NetworkInformation.Ping sqlPing = new System.Net.NetworkInformation.Ping())
                {
                    string[] aryProvider = oConnectionString.Split(new char[1] { ';' });
                    foreach (string item in aryProvider)
                    {
                        if (item.ToLower().Contains("server="))
                        {
                            theServer = item.ToLower().Replace("server=", "");
                        }
                    }
                    aryProvider = null;
                    try
                    {
                        System.Net.NetworkInformation.PingReply sqlReply = sqlPing.Send(theServer);
                        return new SqlConnection(oConnectionString);
                    }
                    catch
                    {
                        /* DEBUG */
                        return null;
                    }
                }
            }
            catch
            {
                /* DEBUG */
                return null;
            }
        }

        public void UpdateDatabase(string dGroup, string dcMembers, string dnMembers, string drMembers)
        {
            try
            {
                // Log
                EventLog.WriteEntry("SGM", "Updating Database (Group Change: " + dGroup + ")...", EventLogEntryType.Information);

                // Open SQL Connection
                using (SqlConnection dbConn = SqlConnect())
                {
                    // Open Database Connection
                    dbConn.Open();

                    // Insert new record
                    try
                    {
                        SqlCommand sCmd = new SqlCommand("INSERT INTO sgmTracking (tGroupname, tMembers, tNewMembers, tRemMembers, tTimeStamp) VALUES ('" + dGroup + "','" + dcMembers + "','" + dnMembers + "','" + drMembers + "','" + DateTime.Now + "')", dbConn);
                        sCmd.ExecuteNonQuery();
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch { /* DEBUG */ }

                    // Close Database Connection
                    dbConn.Close();
                }
            }
            catch { /* DEBUG */ }
        }

        private void UpdateDatabase(string dGroup, string dMembers)
        {
            try
            {
                // Log
                EventLog.WriteEntry("SGM", "Updating Database (Group Added: " + dGroup + ")...", EventLogEntryType.Information);

                // Open SQL Connection
                using (SqlConnection dbConn = SqlConnect())
                {
                    // Open Database Connection
                    dbConn.Open();

                    // Insert new record
                    try
                    {
                        SqlCommand sCmd = new SqlCommand("INSERT INTO sgmTracking (tGroupname, tMembers, tTimeStamp) VALUES ('" + dGroup + "','" + dMembers + "','" + DateTime.Now + "')", dbConn);
                        sCmd.ExecuteNonQuery();
                        sCmd.Dispose();
                        sCmd = null;
                    }
                    catch { /* DEBUG */ }

                    // Close Database Connection
                    dbConn.Close();
                }
            }
            catch { /* DEBUG */ }
        }

        public void Set(string oSetting, string oValue)
        {
            try
            {
                iSettings.Add(oSetting, oValue);
            }
            catch { /* DEBUG */ }
        }

        public string Get(string oSetting)
        {
            string getResult = null;

            try
            {
                getResult = iSettings[oSetting];
            }
            catch { /* DEBUG */ }

            return getResult;
        }

        // Implement IDisposable.
        // Do not make this method virtual.
        // A derived class should not be able to override this method.
        public void Dispose()
        {
            Dispose(true);
            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }

        // Dispose(bool disposing) executes in two distinct scenarios.
        // If disposing equals true, the method has been called directly
        // or indirectly by a user's code. Managed and unmanaged resources
        // can be disposed.
        // If disposing equals false, the method has been called by the
        // runtime from inside the finalizer and you should not reference
        // other objects. Only unmanaged resources can be disposed.
        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    // Dispose managed resources.
                    component.Dispose();
                }

                // Call the appropriate methods to clean up
                // unmanaged resources here.
                // If disposing is false,
                // only the following code is executed.
                CloseHandle(handle);
                handle = IntPtr.Zero;

                // Note disposing has been done.
                disposed = true;

            }
        }

        // Use interop to call the method necessary
        // to clean up the unmanaged resource.
        [System.Runtime.InteropServices.DllImport("Kernel32")]
        private extern static Boolean CloseHandle(IntPtr handle);
    }
}
