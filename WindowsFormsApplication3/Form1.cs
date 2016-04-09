using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Lync.Model;
using System.Net;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Timers;
using System.Collections.Specialized;

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            SetArtistText(_curArtist);
            SetTitleText(_curTitle);
            SetStatusText(_curStatus);
        }

        private bool _isRunning = false;
        private String _curStatus = null;
        private String _curArtist = null;
        private String _curTitle = null;
        private String _initStatus = null;
        private LyncClient _LyncClient = null;
        private String _sonosHost = null;
        private NameValueCollection _appSettings;
        private System.Timers.Timer _connectionTimer = null;

        #region Form Events
        private void button1_Click(object sender, EventArgs e)
        {
            _appSettings = System.Configuration.ConfigurationManager.AppSettings;
            _sonosHost = _appSettings["sonosIP"];
            _LyncClient = LyncClient.GetClient();
            _initStatus = PollStatus();
            _connectionTimer = new System.Timers.Timer(5000);
            _connectionTimer.Elapsed += new ElapsedEventHandler(this.TickTock);
            _connectionTimer.Enabled = true;
            _connectionTimer.Start();
            this.UpdateRunningState();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _connectionTimer.Stop();
            MessageBox.Show("init: " + _initStatus);
            UpdateStatus(_initStatus);
            this.UpdateRunningState();
        }
        #endregion

        #region Lync Functions
        private String PollStatus() 
        {
            return _LyncClient.Self.Contact.GetContactInformation(ContactInformationType.LocationName).ToString();
        }
        private void UpdateStatus(String status)
        {
            Dictionary<PublishableContactInformationType, object> stuff = new Dictionary<PublishableContactInformationType, object>();
            stuff.Add(PublishableContactInformationType.LocationName, status);
            _LyncClient.Self.BeginPublishContactInformation(stuff,
                (ar) =>
                {
                    _LyncClient.Self.EndPublishContactInformation(ar);
                }
                ,
            null);
            this._curStatus = status;
        }
        #endregion

        #region SONOS Functions
        private bool PollSonos()
        {
            bool retVal = true;
            String service = "urn:schemas-upnp-org:service:AVTransport:1";
            String action = "GetPositionInfo";

            // Build the headers
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create("http://" + _sonosHost + "/MediaRenderer/AVTransport/Control");
            try
            {
                req.Headers.Add("SOAPACTION", service + "#" + action);
                req.ContentType = "text/xml;charset=\"utf-8\"";
                req.Method = "POST";
                //req.Timeout = System.Threading.Timeout.Infinite;

                // Write the Request XML
                StreamWriter writer = new StreamWriter(req.GetRequestStream(), new UTF8Encoding(false));
                writer.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?><s:Envelope s:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\" xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\"><s:Body><u:GetPositionInfo xmlns:u=\"urn:schemas-upnp-org:service:AVTransport:1\"><InstanceID>0</InstanceID></u:GetPositionInfo></s:Body></s:Envelope>");
                writer.Close();

                // Get the Response XML
                WebResponse resp = req.GetResponse();
                XmlReader r = new XmlTextReader(resp.GetResponseStream());
                bool nextNode = false;
                String metadata = null;
                while (r.Read())
                {
                    switch (r.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (r.Name == "TrackMetaData")
                                nextNode = true;
                            break;
                        case XmlNodeType.Text:
                            if (nextNode)
                            {
                                metadata = r.Value;
                                nextNode = false;
                            }
                            break;
                    }
                }
                Regex regex = new Regex(".*<dc:creator>(.*)</dc:creator>.*");
                Match match = regex.Match(metadata);
                if (match.Success)
                    _curArtist = System.Web.HttpUtility.HtmlDecode(match.Groups[1].Value);
                else
                {
                    _curArtist = "";
                    retVal = false;
                }
                regex = new Regex(".*<dc:title>(.*)</dc:title>.*");
                match = regex.Match(metadata);
                if (match.Success)
                    _curTitle = System.Web.HttpUtility.HtmlDecode(match.Groups[1].Value);
                else
                {
                    _curTitle = "";
                    retVal = false;
                }
                SetMetadataText(metadata);
            }
            catch (Exception e)
            {
                SetMetadataText("[" + System.DateTime.Now.ToShortTimeString() + "] " + e.Message);
            }
            return retVal;
        }
        private void PollSonosFilepath() 
        {
            String metadata = txtMetadata.Text;
            Regex regex = new Regex(".*<res.*>(.*)</res>.*");
            Match match = regex.Match(metadata);
            String filepath = System.Web.HttpUtility.HtmlDecode(match.Groups[1].Value);
            regex = new Regex(".*MUSIC.*/(.*).mp3");
            match = regex.Match(filepath);
            if (match.Success)
            {
                String[] temp = System.Web.HttpUtility.HtmlDecode(match.Groups[1].Value).Split('-');
//                MessageBox.Show(String.Format("{0}::{1}", temp[0], temp[1]));
                _curArtist = temp[0].Trim();
                _curTitle = temp[1].Trim();
            }
            return;
        }
        #endregion

        #region Utility Functions
        delegate void SetMetadataTextCallback(string text);
        private void SetMetadataText(String newtext)
        {
            if (this.txtMetadata.InvokeRequired)
            {
                SetMetadataTextCallback d = new SetMetadataTextCallback(SetMetadataText);
                this.Invoke(d, new object[] { newtext });
            }
            else
            {
                this.txtMetadata.Text = newtext;
            }
        }
        delegate void SetArtistTextCallback(string text);
        private void SetArtistText(String newtext)
        {
            if (this.txtArtist.InvokeRequired)
            {
                SetArtistTextCallback d = new SetArtistTextCallback(SetArtistText);
                this.Invoke(d, new object[] { newtext });
            }
            else
            {
                this.txtArtist.Text = newtext;
            }
        }
        delegate void SetTitleTextCallback(string text);
        private void SetTitleText(String newtext)
        {
            if (this.txtTitle.InvokeRequired)
            {
                SetTitleTextCallback d = new SetTitleTextCallback(SetTitleText);
                this.Invoke(d, new object[] { newtext });
            }
            else
            {
                this.txtTitle.Text = newtext;
            }
        }
        delegate void SetStatusTextCallback(string text);
        private void SetStatusText(String newtext)
        {
            if (this.txtStatus.InvokeRequired)
            {
                SetStatusTextCallback d = new SetStatusTextCallback(SetStatusText);
                this.Invoke(d, new object[] { newtext });
            }
            else
            {
                this.txtStatus.Text = newtext;
            }
        }
        private void TickTock(object sender, ElapsedEventArgs e)
        {
            if (!PollSonos())
                PollSonosFilepath();
            if (PollStatus() != _appSettings["message"].Replace("$artist", _curArtist).Replace("$title", _curTitle))
            {
                UpdateStatus(_appSettings["message"].Replace("$artist", _curArtist).Replace("$title", _curTitle));
            }
            SetArtistText(_curArtist);
            SetTitleText(_curTitle);
            SetStatusText(_curStatus);
        }
        private void UpdateRunningState() 
        {
            btnStart.Enabled = !btnStart.Enabled;
            btnStop.Enabled = !btnStop.Enabled;
            _isRunning = !_isRunning;
            return;
        }
        #endregion

    }
}
