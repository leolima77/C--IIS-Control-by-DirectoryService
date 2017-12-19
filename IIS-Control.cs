using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Net;

public class IIS
{
    #region private properties
    private string _serverName;
    private string _websiteName;
    #endregion

    #region public methods
    public IIS(string serverName, string websiteName)
    {
        _serverName = serverName;
        _websiteName = websiteName;
    }

    public int GetWebSiteId()
    {
        string result = "-1";

        DirectoryEntry w3svc = new DirectoryEntry(string.Format("IIS://" + _serverName + "/w3svc"));

        //Para problemas: https://stackoverflow.com/questions/1722398/error-0x80005000-and-directoryservices
        //Mais problemas https://forums.iis.net/post/2018190.aspx
        foreach (DirectoryEntry site in w3svc.Children)
        {
            if (site.Properties["ServerComment"] != null)
            {
                if (site.Properties["ServerComment"].Value != null)
                {
                    if (string.Compare(site.Properties["ServerComment"].Value.ToString(),
                                            _websiteName, true) == 0)
                    {
                        result = site.Name;
                        break;
                    }
                }
            }
        }

        return Convert.ToInt32(result);
    }

    public bool AddHostHeader(string hostHeader)
    {
        if (hostHeader.IndexOf(':') < 0)
            return false;

        using (DirectoryEntry site = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId()))
        {
            try
            {
                PropertyValueCollection serverBindings = site.Properties["ServerBindings"];

                if (ExistHostHeader(hostHeader))
                    return false;

                serverBindings.Add(hostHeader);

                Object[] newList = new Object[serverBindings.Count];
                serverBindings.CopyTo(newList, 0);

                site.Properties["ServerBindings"].Value = newList;

                site.CommitChanges();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    public bool RemoveHostHeader(string hostHeader)
    {
        using (DirectoryEntry site = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId()))
        {
            try
            {
                if (ExistHostHeader(hostHeader))
                {
                    site.Properties["ServerBindings"].Remove(hostHeader);
                    site.CommitChanges();
                    return true;
                }
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
    }

    public List<string> ListHostHeader()
    {
        using (DirectoryEntry site = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId()))
        {
            PropertyValueCollection serverBindings = site.Properties["ServerBindings"];

            List<string> retorno = new List<string>();
            for (int i = 0; i < serverBindings.Count; i++)
                retorno.Add(serverBindings[i].ToString());

            return retorno;
        }
    }

    public bool ExistHostHeader(string hostHeader)
    {
        using (DirectoryEntry site = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId()))
        {
            if (site.Properties["ServerBindings"].Contains(hostHeader))
                return true;

            return false;
        }
    }

    public List<string> ListSubsVirtualDir()
    {
        using (DirectoryEntry websiteEntry = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId() + "/root"))
        {
            List<string> retorno = new List<string>();
            foreach (DirectoryEntry entry in websiteEntry.Children)
            {
                if (entry.SchemaClassName == "IIsWebVirtualDir")
                    retorno.Add(entry.Name);
            }

            return retorno;
        }
    }

    public bool CreateSubVirtualDir(string vDirName, string physicalPath)
    {
        try
        {
            string metabase = "IIS://" + _serverName + "/w3svc/" + GetWebSiteId() + "/root";

            using (DirectoryEntry site = new DirectoryEntry(metabase))
            {
                string className = site.SchemaClassName.ToString();
                if ((className.EndsWith("Server")) || (className.EndsWith("VirtualDir")))
                {
                    DirectoryEntries vdirs = site.Children;
                    DirectoryEntry newVDir = vdirs.Add(vDirName, (className.Replace("Service", "VirtualDir")));
                    newVDir.Properties["Path"][0] = physicalPath;
                    newVDir.Properties["AccessScript"][0] = true;
                    // These properties are necessary for an application to be created.
                    newVDir.Properties["AppFriendlyName"][0] = vDirName;
                    newVDir.Properties["AppIsolated"][0] = "1";
                    newVDir.Properties["AppRoot"][0] = "/LM" + metabase.Substring(metabase.IndexOf("/", ("IIS://".Length)));

                    newVDir.CommitChanges();

                    return true;
                }
                else
                    return false;
            }
        }
        catch (Exception ex)
        {
            return false;
        }
    }

    public bool RemoveSubsVirtualDir(string subVirtualName)
    {
        // get directory service
        using (DirectoryEntry websiteEntry = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId() + "/root"))
        {
            var diretorios = ListSubsVirtualDirDirectoryEntry();
            foreach (DirectoryEntry diretorio in diretorios)
            {
                if (diretorio.Name.ToLower() == subVirtualName.ToLower())
                    try
                    {
                        websiteEntry.Children.Remove(diretorio);
                        websiteEntry.CommitChanges();

                        return true;
                    }
                    catch
                    {
                        return false;
                    }
            }

            return false;
        }
    }

    public string OpenAppPool(string name)
    {
        string connectStr = "IIS://" + _serverName + "/w3svc/AppPools/";
        connectStr += name;

        if (ExistAppPool(name) == false)
            return null;

        using (DirectoryEntry entry = new DirectoryEntry(connectStr))
        {
            return entry.Name;
        }
    }

    public bool ExistAppPool(string name)
    {
        using (DirectoryEntry Service = new DirectoryEntry("IIS://" + _serverName + "/w3svc/AppPools"))
        {
            foreach (DirectoryEntry entry in Service.Children)
            {
                if (entry.Name.Trim().ToLower() == name.Trim().ToLower())
                    return true;
            }
            return false;
        }
    }
    #endregion

    #region private methods
    private DirectoryEntry OpenWebsite()
    {
        // get directory service
        using (DirectoryEntry Services = new DirectoryEntry("IIS://" + _serverName + "/w3svc"))
        {
            IEnumerator ie = Services.Children.GetEnumerator();
            DirectoryEntry Server = null;

            // find iis website
            while (ie.MoveNext())
            {
                Server = (DirectoryEntry)ie.Current;
                if (Server.SchemaClassName == "IIsWebServer")
                    // "ServerComment" means name
                    if (Server.Properties["ServerComment"][0].ToString() == _websiteName)
                        return Server;
            }

            return null;
        }
    }

    private DirectoryEntry OpenSubsVirtualDir(string subVirtualName)
    {
        // get directory service
        using (DirectoryEntry websiteEntry = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId() + "/root"))
        {
            foreach (DirectoryEntry entry in websiteEntry.Children)
            {
                if (entry.SchemaClassName == "IIsWebVirtualDir" && entry.Name.ToLower() == subVirtualName.ToLower())
                    return entry;
                else
                    return null;
            }

            return null;
        }
    }

    private List<DirectoryEntry> ListSubsVirtualDirDirectoryEntry()
    {
        using (DirectoryEntry websiteEntry = new DirectoryEntry("IIS://" + _serverName + "/w3svc/" + GetWebSiteId() + "/root"))
        {
            List<DirectoryEntry> retorno = new List<DirectoryEntry>();
            foreach (DirectoryEntry entry in websiteEntry.Children)
            {
                if (entry.SchemaClassName == "IIsWebVirtualDir")
                    retorno.Add(entry);
            }

            return retorno;
        }
    }
    #endregion
}