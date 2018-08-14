using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Newtonsoft.Json;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParagraphFinder
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string userName = Environment.UserName;
            string fileLocationStart = "c:\\Users\\" + userName + "\\Desktop\\";

            Console.WriteLine("Please enter a keyword:");
            string keyword = Console.ReadLine();

            Console.WriteLine("\nPlease select a file:");

            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Title = "Open File";
            fbd.Filter = "WORD (.docx, or .doc,)|*.docx;*.doc";
            fbd.InitialDirectory = fileLocationStart;

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string ext = System.IO.Path.GetExtension(fbd.FileName);
                string convText = "";

                //Convert file into text
                try
                {
                    //Using Spire office library instead of interop because interop is slow and Microsoft does not currently recommend,
                    //and does not support, Automation of Microsoft Office applications from any unattended non-interactive client application or component
                    using (var stream1 = new MemoryStream())
                    {
                        MemoryStream txtStream = new MemoryStream();
                        Document document = new Document();
                        document.LoadFromFile(fbd.FileName);
                        document.SaveToStream(txtStream, FileFormat.Txt);
                        txtStream.Position = 0;

                        StreamReader reader = new StreamReader(txtStream);
                        string readText = reader.ReadToEnd();

                        //Remove watermark for spire
                        readText = readText.Replace("Evaluation Warning: The document was created with Spire.Doc for .NET.", "");
                        convText = readText;
                    }
                }
                catch
                {
                    MessageBox.Show(fbd.FileName + " cannot be opened! Skipping this file.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                string postData = "[";

                List<string> paragraphs = new List<string>();

                int count = 0;
                foreach (string paragraph in convText.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (String.IsNullOrWhiteSpace(paragraph) == false)
                    {
                        paragraphs.Add(paragraph);
                        List<Char> builder = new List<char>();
                        //Used to fix if there are multiple newlines in a row
                        bool isNewLine = true;

                        //Remove special characters which would need to be escaped for JSON and creates new string using builder var
                        for (int i = 0; i < paragraph.Length; i++)
                        {
                            if (paragraph[i] == '\t')
                            {
                                builder.Add(' ');
                            }
                            else if (paragraph[i] == char.MinValue)
                            {
                                builder.Add(' ');
                            }
                            else if (paragraph[i] == '\\')
                            {
                                builder.Add('\\');
                                builder.Add('\\');
                            }
                            else if ((paragraph[i] == '\n' || paragraph[i] == '\r') && isNewLine == false)
                            {
                                if (paragraph[i - 1] == '.' || paragraph[i - 1] == ':' || paragraph[i - 1] == ',')
                                {
                                    builder.Add(' ');
                                }
                                else if (paragraph[i - 1] != ' ')
                                {
                                    builder.Add('.');
                                    builder.Add(' ');
                                }
                                isNewLine = true;
                            }
                            else if (paragraph[i] != '\n' && paragraph[i] != '\r')
                            {
                                isNewLine = false;
                                //If '"' is already escaped ignore
                                if (paragraph[i] == '"' && paragraph[i - 1] != '\\')
                                {
                                    //Adds a single '\' before the '"'
                                    builder.Add('\\');
                                    builder.Add('"');
                                }
                                else
                                {
                                    builder.Add(paragraph[i]);
                                }
                            }
                        }
                        string newConvText = new string(builder.ToArray());

                        postData += "[{\"term\": \"" + keyword + "\"},{\"text\": \"" + newConvText + "\"}],";

                        count++;
                    }
                }
                postData = postData.Remove(postData.Length - 1, 1) + "]";

                //API Request to cortical.io to compare text taken from document with a keyword the user provided
                WebRequest webRequest = WebRequest.Create("http://api.cortical.io:80/rest/compare/bulk?retina_name=en_associative");
                webRequest.Method = "POST";
                webRequest.Headers["api-key"] = "bb355cc0-5873-11e8-9172-3ff24e827f76";
                webRequest.ContentType = "application/json";
                //Send request with postData string as the body
                using (var streamWriter = new StreamWriter(webRequest.GetRequestStream()))
                {
                    streamWriter.Write(postData);
                    streamWriter.Flush();
                    streamWriter.Close();
                }
                string result = "";
                //Recieve response from cortical.io API
                try
                {
                    WebResponse webResp = webRequest.GetResponse();
                    using (var streamReader = new StreamReader(webResp.GetResponseStream()))
                    {
                        result = streamReader.ReadToEnd();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("\nCannot connect to cortical.io API. Aborting!\n\nError: " + ex.Message);
                    Console.ReadLine();
                    return;
                }

                //Formats return string as JSON
                dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(result);

                List<KeyValuePair<double, int>> cosineNum = new List<KeyValuePair<double, int>>();
                List<KeyValuePair<double, string>> cosineParagraph = new List<KeyValuePair<double, string>>();

                //Calculates match percent for each return object which correlates to each resume
                for (int i = 0; i < jsonObj.Count; i++)
                {
                    double cosineSim = Math.Round((double)jsonObj[i].cosineSimilarity, 3);

                    cosineNum.Add(new KeyValuePair<double, int>(cosineSim, i));
                    cosineParagraph.Add(new KeyValuePair<double, string>(cosineSim, paragraphs[i]));
                }
                cosineNum = cosineNum.OrderByDescending(x => x.Key).ToList();
                cosineParagraph = cosineParagraph.OrderByDescending(x => x.Key).ToList();

                Console.WriteLine("\nResults:\n");
                for (int i = 0; i < cosineNum.Count; i++)
                {
                    if (cosineNum[i].Key >= 0.25)
                    {
                        Console.WriteLine(i + 1 + ".) " + cosineParagraph[i].Value + "\n" + cosineNum[i].Key + " Cosine Similarity\n");
                    }
                }
                Console.ReadLine();
            }
        }
    }
}
