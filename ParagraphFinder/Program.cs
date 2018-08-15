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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParagraphFinder
{
    class Program
    {
        //To allow Open File Dialogue
        [STAThread]
        static void Main(string[] args)
        {
            List<string>[] library = null;
            bool isLibEnabled = true;

            try
            {
                //Read LibraryCSV file
                var lines = File.ReadLines("LibraryCSV.txt");
                int lineCount = File.ReadLines("LibraryCSV.txt").Count();

                library = new List<string>[lineCount];
                int count2 = 0;

                //Go through each line (categories) of the file
                foreach (var line in lines)
                {
                    //Load array with line contents
                    var values = line.Split(',');
                    //Initialize new list in array of lists
                    library[count2] = new List<string>();
                    //Add contents of array to list
                    foreach (var item in values)
                    {
                        library[count2].Add(item);
                    }
                    count2++;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Lib file cannot be found or is not formatted in proper csv format! User will be unable to use library functionality and only cortical.io functionality.\nError Message:\n\n" + e.Message + "\n\n");
                isLibEnabled = false;
            }

            //Get starting location for Open File Dialogue
            string userName = Environment.UserName;
            string fileLocationStart = "c:\\Users\\" + userName + "\\Desktop\\";

            string keyword = "";
            //Continously ask user for keyword until a proper keyword is entered
            do
            {
                Console.WriteLine("Please enter a keyword:");
                keyword = Console.ReadLine();

            }
            while (keyword == "");

            int inLibNum = -1;
            //Check if using library
            if (isLibEnabled == true)
            {
                //Find which category from the library the keyword is in
                for (int i = 0; i < library.Length; i++)
                {
                    if (library[i].Contains(keyword, StringComparer.OrdinalIgnoreCase))
                    {
                        inLibNum = i;
                        break;
                    }
                }
                if (inLibNum != -1)
                {
                    Console.WriteLine("Using library! From row " + (inLibNum + 1));
                }
            }

            Console.WriteLine("\nPlease select a file:");

            //Request file
            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Title = "Open File";
            fbd.Filter = "WORD (.docx, or .doc,)|*.docx;*.doc";
            fbd.InitialDirectory = fileLocationStart;

            //File successfully found
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
                    MessageBox.Show(fbd.FileName + " cannot be opened!\nPress any key to end program.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Console.ReadKey();
                    return;
                }

                string postData = "[";

                List<string> paragraphs = new List<string>();
                List<KeyValuePair<double, int>> matchNum = new List<KeyValuePair<double, int>>();
                List<KeyValuePair<double, string>> matchParagraph = new List<KeyValuePair<double, string>>();

                int count = 0;
                //Loop through each paragraph in the file
                foreach (string paragraph in convText.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    //Not a empty line
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

                        //Build api post request for cortical
                        postData += "[{\"term\": \"" + keyword + "\"},{\"text\": \"" + newConvText + "\"}],";

                        if (inLibNum != -1)
                        {
                            int numExactMatches = 0;
                            int numCategoryMatches = 0;

                            for (int i = 0; i < library[inLibNum].Count; i++)
                            {
                                //Paragraph contains the keyword
                                if (newConvText.IndexOf(library[inLibNum].ElementAt(i), StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    //Count occurences
                                    int occurences = Regex.Matches(newConvText, @"\b" + library[inLibNum].ElementAt(i), RegexOptions.IgnoreCase).Count;

                                    //Exact match
                                    if (String.Equals(keyword, library[inLibNum].ElementAt(i), StringComparison.OrdinalIgnoreCase))
                                    {
                                        numExactMatches = numExactMatches + occurences;
                                    }
                                    //Categorical match
                                    else
                                    {
                                        numCategoryMatches = numCategoryMatches + occurences;
                                    }
                                }
                            }
                            //Calculate match percent
                            double totalMatchScore = 0;
                            totalMatchScore = (numExactMatches + ((double)numCategoryMatches / 10)) * 10;

                            matchNum.Add(new KeyValuePair<double, int>(totalMatchScore, count));
                            matchParagraph.Add(new KeyValuePair<double, string>(totalMatchScore, newConvText));
                        }
                        count++;
                    }
                }
                //Remove trailing ',' and replace with ']' closing the post request
                postData = postData.Remove(postData.Length - 1, 1) + "]";

                //Use cortical.io
                if (inLibNum == -1)
                {
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
                        Console.WriteLine("\nCannot connect to cortical.io API. Aborting!\n\nError: " + ex.Message + "\nPress any key to end program.");
                        Console.ReadKey();
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

                    bool resultsFound = false;

                    Console.WriteLine("\nResults:\n");
                    for (int i = 0; i < cosineNum.Count; i++)
                    {
                        //Only display results with cosine of 0.25 or greater
                        if (cosineNum[i].Key >= 0.25)
                        {
                            resultsFound = true;
                            Console.WriteLine(i + 1 + ".) " + cosineParagraph[i].Value + "\n" + cosineNum[i].Key + " Cosine Similarity\n");
                        }
                    }
                    //No results found
                    if (resultsFound == false)
                    {
                        Console.WriteLine("No results can be found with a cosine similarity greater or equal to 0.25.");
                    }
                }
                //Use library
                else
                {
                    bool resultsFound = false;

                    //Order results
                    matchNum = matchNum.OrderByDescending(x => x.Key).ToList();
                    matchParagraph = matchParagraph.OrderByDescending(x => x.Key).ToList();

                    Console.WriteLine("\nResults:\n");
                    for (int i = 0; i < matchNum.Count; i++)
                    {
                        //Only show results with atleast 1% match
                        if (matchNum[i].Key >= 1)
                        {
                            resultsFound = true;
                            Console.WriteLine(i + 1 + ".) " + matchParagraph[i].Value + "\n" + matchNum[i].Key + "% match\n");
                        }
                    }
                    //No results found
                    if (resultsFound == false)
                    {
                        Console.WriteLine("No results can be found with a match score greater or equal to 1%.");
                    }
                }
                Console.WriteLine("\nPress any key to end program.");
                Console.ReadKey();
            }
            //User closed File Selection Dialogue
            else
            {
                Console.WriteLine("\nFile Not Selected! Press any key to end program.");
                Console.ReadKey();
            }
        }
    }
}
