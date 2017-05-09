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

namespace WindowsFormsApp1
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public static bool isRoot = false;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sitecoll = DetectSiteCollectionURL(textBox1.Text);
            string list = DetectListNameURL(sitecoll, textBox1.Text);
            string FileName = DetectFileName(textBox1.Text);

            MessageBox.Show(sitecoll);
            MessageBox.Show(list);
            MessageBox.Show(FileName);

            connectTOSP(sitecoll, list, FileName);
        }

        private string DetectFileName(string url)
        {
            string[] test = url.Split('/');

            return System.Web.HttpUtility.UrlDecode(test[test.Length - 1]);

        }
        private string DetectSiteCollectionURL(string url)
        {

            if (url.ToLower().Contains(textBox5.Text.ToLower()))
            {
                //contains managed path
                isRoot = false;
                return DetectNonRootSiteCollectionURL(url);

            }
            else
            {
                //contains non-root site collection
                isRoot = true;
                return DetectRootSiteCollectionURL(url);

            }

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
        private void connectTOSP(string siteurl, string listname, string filename)
        {
            MessageBox.Show(DetectList(siteurl, listname, filename));
            
        }

        private string DetectList(string siteurl, string listname, string filename)
        {
            string status = "";
            try
            {
                //try if its a list
                status= TryGettingList(siteurl, listname, filename);
            }
            catch (Exception)
            {

                //try if its a library
                status= TryGettingLibrary(siteurl, listname, filename);
            }

            return status;

        }

        private string TryGettingList(string siteurl, string listname, string filename)
        {
            string status = "";
            using (ClientContext ctx = new ClientContext(siteurl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(textBox2.Text, ConvertToSecureString(textBox3.Text));

                ListCollection AllList = ctx.Web.Lists;
                List Mylist = ctx.Web.GetList(listname);
                ctx.Load(Mylist);

                ctx.ExecuteQuery();


                status = $"List {Mylist.Title}{Environment.NewLine}";
                
            }
            return status;
        }
        
        private string TryGettingLibrary(string siteurl, string listname, string filename)
        {
            string status = "";
            using (ClientContext ctx = new ClientContext(siteurl))
                {
                    ctx.Credentials = new SharePointOnlineCredentials(textBox2.Text, ConvertToSecureString(textBox3.Text));

                    ListCollection AllList = ctx.Web.Lists;
                    ctx.Load(AllList, lst => lst.Include(l => l.Title,
                                                         l => l.DefaultViewUrl)
                                                 .Where(l => l.IsSystemList != true));


                    ctx.Load(AllList);

                    ctx.ExecuteQuery();

                    foreach (List lst in AllList)
                    {
                        if (lst.DefaultViewUrl.ToString().Contains(listname))
                        {
                            status =  $"Document library -{lst.Title}{Environment.NewLine}";
                        }

                    }
                    
                }
            return status;

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
    }
}
