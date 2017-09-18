using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Amazon.Lex;
using Amazon;
using System.Speech.Synthesis;
using System.IO;
using System.Configuration;

namespace Zeus
{
    public partial class Zeus : Form
    {
        public static AmazonLexClient lexClient;
        SpeechSynthesizer _synthesizer = new SpeechSynthesizer();
        
        public Zeus()
        {
            InitializeComponent();
            MaximizeBox = false;
            
        }
        Bitmap[] bmps = null;
        int ind = 0;
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                GetAmazonLexClient();
                SetBotFaceImages();
                RegisterSpeakSynthesizerEvents();
            }
            catch (Exception ex)
            {

                LogEvent("ERROR : "+ex.Message);
            }       
                       
        }

        private void RegisterSpeakSynthesizerEvents()
        {
            _synthesizer.SpeakProgress += new EventHandler<SpeakProgressEventArgs>(synthesizer_SpeakProgress);

            _synthesizer.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(synthesizer_SpeakCompleted);
        }

        private void SetBotFaceImages()
        {
            string dname = ".\\Assets\\";
            string[] fnames = Directory.GetFiles(dname, "*.jpg");
            bmps = new Bitmap[fnames.Length];

            for (int i = 0; i < fnames.Length; i++)
            {
                bmps[i] = Bitmap.FromFile(fnames[i]) as Bitmap;
                bmps[i].MakeTransparent();
            }

            pictureBox1.Image = bmps[ind % bmps.Length];
            ind++;
        }
        
        public static AmazonLexClient GetAmazonLexClient()
        {
            string AWSKey = ConfigurationSettings.AppSettings["AWSKey"];

            string AWSSecKey = ConfigurationSettings.AppSettings["AWSSecKey"];
            
            lexClient = new AmazonLexClient(AWSKey, AWSSecKey, RegionEndpoint.USEast1);
            return lexClient;
        }

        public static string LexRequest(string userMessage)
        {
            try
            {
                var lexResponse = lexClient.PostText(
               new Amazon.Lex.Model.PostTextRequest()
               {
                    //BotAlias = "skypeChitChat",
                    //BotName = "chitChat",
                    //InputText = userMessage,
                    //UserId = "Anupam"
                    BotAlias = "Zeus",
                   BotName = "Zeus",
                   InputText = userMessage,
                   UserId = "Anupam"
               });
                return lexResponse.Message.Replace("*",string.Empty);
            }
            catch (Exception ex)
            {                
                return "Keyword not found.";
            }
            
           
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text.Trim() != string.Empty)
            {
                //richTextBox2.Text = richTextBox1.Text;

                richTextBox2.Text = LexRequest(richTextBox1.Text);
                var currentSpeaker = _synthesizer.GetCurrentlySpokenPrompt();
                if(currentSpeaker!=null)
                _synthesizer.SpeakAsyncCancel(currentSpeaker);

                speak();
                richTextBox1.Text = string.Empty;
            }
            else
            {
                richTextBox2.Text = "Oh! You forgot to enter text message.";
                speak();
            }
        }

        private void speak()
        {
            _synthesizer.SpeakAsync(richTextBox2.Text);
         
        }

        void synthesizer_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            pictureBox1.Image = bmps[0];
            richTextBox1.Text = string.Empty;
        }

        void synthesizer_SpeakProgress(object sender, SpeakProgressEventArgs e)
        {
            pictureBox1.Image = bmps[ind % bmps.Length];
            ind++;
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSend_Click(this, new EventArgs());
            }
        }

        private void LogEvent(string msg)
        {
            File.AppendAllText(@"C:\Users\admin\source\repos\ConsoleApp1\Zeus\Logs\ZeusLog.txt", msg + Environment.NewLine);
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }
    }
}
