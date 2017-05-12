using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Net;

namespace WindowsFormsApp1
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public static bool isRoot = false;
        public static ClientContext GlobalCtx = null;
        public static bool isSPO = true;
        public static bool isSpecialFolders = false;
        public string list = "";
        public string FileName = "";
        public int SpecialFolderType = 0;
        public string CSVFileLocation = "";
        public string[] AllURL = null;
        public string GlobalManagedPath ="";
        public string OutputFileLocation = "";

        public enum SpecialFolderTypeEnum {webpart,masterpage,listtemplate,solution};

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (string url in AllURL)
            {
                LogtoFile(url + ";");
                DoWork(url);
            }
                      
        }

        private void DoWork(string url)
        {
            GlobalManagedPath = textBox5.Text;
            string sitecoll = DetectSiteCollectionURL(url);
            isSPO = checkBox1.Checked;
            

            if (!isSpecialFolders)
            {
                list = System.Web.HttpUtility.UrlDecode(DetectListNameURL(sitecoll, url));
                FileName = DetectFileName(url);
                LogtoFile(connectTOSP(sitecoll, list, FileName));
            }

        }

        private void LogtoFile(string msg)
        {
            System.IO.File.AppendAllText(OutputFileLocation,msg);
        }

        private string DetectFileName(string url)
        {
            string[] test = url.Split('/');

            return System.Web.HttpUtility.UrlDecode(test[test.Length - 1]);

        }

        private string DetectSiteCollectionURL(string url)
        {
            if (url.ToLower().Contains("_catalogs"))
            {
                isSpecialFolders = true;
                
                //masterpages and other special folders
                return DetectSpecialFolderFiles(url, DetectFileName(url));
            }
            else if (url.ToLower().Contains(textBox5.Text.ToLower()))
            {
                //contains managed path
                isSpecialFolders = false;
                isRoot = false;
                return DetectNonRootSiteCollectionURL(url);

            }
            else 
            {
                //contains non-root site collection
                isRoot = true;
                isSpecialFolders = false;
                return DetectRootSiteCollectionURL(url);

            }

        }

        private string DetectSpecialFolderFiles(string url,string filename)
        {
            Uri SiteCollUri = new Uri(url);

            //eg - csv gives you this - https://mstnd550949.sharepoint.com/sites/contoso/_catalogs/wp/BusinessDataFilter.dwp

            //the below gives you this - https://mstnd550949.sharepoint.com/
            string baseurl = $"{SiteCollUri.Scheme}://{SiteCollUri.Host}/{GlobalManagedPath}";

            // the below gives you - sites/contoso/_catalogs/wp/BusinessDataFilter.dwp
            string[] temp = url.Substring(baseurl.Length).Split('/');
            string sitecollname = temp[1];
            string sitecollurl = $"{baseurl}/{sitecollname}/";

            


            if (url.ToLower().Contains("/_catalogs/wp")) // this is webpart reference
            {
                
                
                handleSpecialFolder(SpecialFolderTypeEnum.webpart,sitecollurl,filename);

            }
            else if (url.ToLower().Contains("/_catalogs/lt")) // this is list template gallary
            {
                
                handleSpecialFolder(SpecialFolderTypeEnum.listtemplate, sitecollurl, filename);
            }
            else if (url.ToLower().Contains("/_catalogs/masterpage")) //this is master page
            {
                handleSpecialFolder(SpecialFolderTypeEnum.masterpage, sitecollurl, filename);
            }
            else if (url.ToLower().Contains("/_catalogs/solution")) // this is solutions gallery
            {
                handleSpecialFolder(SpecialFolderTypeEnum.solution,sitecollurl, filename);
            }

            return "";
            
        }

        private void handleSpecialFolder(SpecialFolderTypeEnum SPType,string url,string filename)
        {
            if (SpecialFolderTypeEnum.webpart == SPType)
            {
                GetItemsOwnerForSpecialFolder(url, "Web Part Gallery", filename);

            }
            else if (SPType == SpecialFolderTypeEnum.listtemplate)
            {
                GetItemsOwnerForSpecialFolder(url, "List Template Gallery", filename);

            }
            else if (SPType == SpecialFolderTypeEnum.masterpage)
            {
                GetItemsOwnerForSpecialFolder(url, "Master Page Gallery", filename);
            }
            else if (SPType == SpecialFolderTypeEnum.solution)
            {
                GetItemsOwnerForSpecialFolder(url, "Solution Gallery", filename);
            }
            else {
                MessageBox.Show("cannot detect the special folder type");
            }
        }

        private string GetItemsOwnerForSpecialFolder(string url, string libraryName, string filename)
        {
            string listname = "";

            using (ClientContext ctx = new ClientContext(url))
            {
                if (isSPO)
                {
                    ctx.Credentials = new SharePointOnlineCredentials(textBox2.Text, ConvertToSecureString(textBox3.Text));
                }
                else
                {
                    NetworkCredential _myCredentials = new NetworkCredential($"{textBox2.Text}", $"{textBox3.Text}");
                    ctx.Credentials = _myCredentials;

                }

                ctx.Load(ctx.Web.Lists, l => l.Include(lst => lst.Title).Where(lst => lst.Title == libraryName));

                ctx.ExecuteQuery();

                List list = null;
                Web web = null;
                ListItemCollection AllItems = null;

                foreach (List l in ctx.Web.Lists)
                {
                    web = ctx.Web;
                    list = l;
                    var q = new CamlQuery() { ViewXml = $"<View><Query><Where><Eq><FieldRef Name='LinkFilename' /><Value Type='Computed'>{filename}</Value></Eq></Where></Query><ViewFields><FieldRef Name='Editor' /><FieldRef Name='Author' /></ViewFields><QueryOptions /></View>" };
                    AllItems = list.GetItems(q);
                    ctx.Load(AllItems);
                    ctx.ExecuteQuery();
                }
                foreach (var item in AllItems)
                {
                    FieldUserValue oValue = item["Author"] as FieldUserValue;
                    MessageBox.Show(oValue.LookupValue);
                    LogtoFile(oValue.LookupValue + Environment.NewLine);
                }
            }
            return listname;
        }

        private string DetectListNameURL(string URL, string rawFullURL)
        {
            if (isRoot)
            {
                return DetectListNameForRootFromURL(URL, rawFullURL);
            }
            else
            {
                return DetectListNameForNonRootFromURL(URL, rawFullURL);
            }



        }

        private string DetectNonRootSiteCollectionURL(string rawURL)
        {

            //only meant for non-root site url

            //remove all the special char to normal chars
            string _rawurl = System.Web.HttpUtility.UrlDecode(rawURL);

            //convert string to url
            Uri _uri = new Uri(_rawurl);

            //get the http part
            string scheme = _uri.Scheme;

            //get the server name
            string server = _uri.Host;

            //get the managed path 'SITES'
            string managedpath = textBox5.Text;
            GlobalManagedPath = managedpath;

            //get the sitecollection name
            string sitecoll = (_uri.Segments)[2].ToString();

            //final url for site collection
            string FullSiteCollURL = $"{scheme}://{server}/{managedpath}/{sitecoll}";

            ////get the library name
            //string lib = (_rawurl.Substring(FullSiteCollURL.Length).Split('/')[0]);

            return FullSiteCollURL;

        }

        private string DetectRootSiteCollectionURL(string rawURL)
        {

            //only meant for rootsite collection

            //remove all the special char to normal chars
            string _rawurl = System.Web.HttpUtility.UrlDecode(rawURL);

            //convert string to url
            Uri _uri = new Uri(_rawurl);

            //get the http part
            string scheme = _uri.Scheme;

            //get the server name
            string server = _uri.Host;


            //get the sitecollection name
            string sitecoll = (_uri.Segments)[2].ToString();

            //final url for site collection
            string FullSiteCollURL = $"{scheme}://{server}";

            return FullSiteCollURL;

        }

        private string DetectListNameForRootFromURL(string SiteCollectionURL, string rawFullURL)
        {
            //get the library name
            return (rawFullURL.Substring(SiteCollectionURL.Length).Split('/')[1]);

        }

        private string DetectListNameForNonRootFromURL(string SiteCollectionURL, string rawFullURL)
        {
            //get the library name
            return (rawFullURL.Substring(SiteCollectionURL.Length).Split('/')[0]);

        }

        private string connectTOSP(string siteurl, string listname, string filename)
        {
          return GetItemsOwner(siteurl,DetectList(siteurl, listname), filename);
            
        }

        private List DetectList(string siteurl, string listname)
        {
            List FinalList = null;
            try
            {
                //try if its a list
                //MessageBox.Show("trying list");
                FinalList = TryGettingList(siteurl, listname);
                
            }
            catch (Exception)
            {

                //try if its a library
                //MessageBox.Show("trying Library");
                FinalList = TryGettingLibrary(siteurl, listname);
            }

            return FinalList;

        }

        private List TryGettingList(string siteurl, string listname)
        {
            List myList = null;
            using (ClientContext ctx = new ClientContext(siteurl))
            {
                if (isSPO)
                {
                    ctx.Credentials = new SharePointOnlineCredentials(textBox2.Text, ConvertToSecureString(textBox3.Text));
                }
                else
                {
                    NetworkCredential _myCredentials = new NetworkCredential($"{textBox2.Text}",$"{textBox3.Text}");
                    ctx.Credentials = _myCredentials;

                }
                GlobalCtx = ctx;
                ListCollection AllList = ctx.Web.Lists;
                List Mylist = ctx.Web.GetList(listname);
                ctx.Load(Mylist);

                ctx.ExecuteQuery();
                
                myList = Mylist;
                
            }
            return myList;
        }

        private string GetItemsOwner(string siteurl,List _list,string filename)
        {
            string strAuthorName = "";
            using (ClientContext ctx = GlobalCtx)
            {
                Web web = ctx.Web;
                ctx.Load(_list);
                ctx.ExecuteQuery();
                List list = _list;

                var q = new CamlQuery() { ViewXml = $"<View><Query><Where><Eq><FieldRef Name='LinkFilename' /><Value Type='Computed'>{filename}</Value></Eq></Where></Query><ViewFields><FieldRef Name='Editor' /><FieldRef Name='Author' /></ViewFields><QueryOptions /></View>" };

                var r = list.GetItems(q);
                ctx.Load(r);
                ctx.ExecuteQuery();

                foreach (var item in r)
                {
                    FieldUserValue oValue = item["Author"] as FieldUserValue;
                    strAuthorName = oValue.LookupValue;
                    //LogtoFile(strAuthorName + Environment.NewLine);
                    
                }


            }

            return strAuthorName;
        }

        private List TryGettingLibrary(string siteurl, string listname)
        {
            List MyList = null;
            try
            {

            using (ClientContext ctx = new ClientContext(siteurl))
                {
                if (isSPO)
                {
                    ctx.Credentials = new SharePointOnlineCredentials(textBox2.Text, ConvertToSecureString(textBox3.Text));
                }
                else
                {
                    NetworkCredential _myCredentials = new NetworkCredential($"{textBox2.Text}", $"{textBox3.Text}");
                    ctx.Credentials = _myCredentials;

                }
                GlobalCtx = ctx;

                    ListCollection AllList = ctx.Web.Lists;
                    ctx.Load(AllList, lst => lst.Include(l => l.Title, l=>l.DefaultViewUrl));
                
                    ctx.Load(AllList);

                    ctx.ExecuteQuery();

                    foreach (List lst in AllList)
                    {
                      
                        if (listname.ToLower().ToString() == (lst.DefaultViewUrl.Split('/'))[3].ToLower().ToString())
                    {
                        //MessageBox.Show("match");
                        MyList = lst;
                    }
                    
                    }
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception trying at library " +ex.Message);
            }
            return MyList;

        }

        private static SecureString ConvertToSecureString(string strPassword)
        {
            var secureStr = new SecureString();
            if (strPassword.Length > 0)
            {
                foreach (var c in strPassword.ToCharArray()) secureStr.AppendChar(c);
            }
            return secureStr;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    AllURL = System.IO.File.ReadAllLines(openFileDialog1.FileName);

                    foreach (string item in AllURL)
                    {
                        textBox4.Text += item + Environment.NewLine;
                    }


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                OutputFileLocation = saveFileDialog1.FileName;
            }
        }
    }
}
