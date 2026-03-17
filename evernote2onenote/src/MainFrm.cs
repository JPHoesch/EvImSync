// Evernote2Onenote - imports Evernote notes to Onenote
// Copyright (C) 2014, 2023 - Stefan Kueng

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using Evernote2Onenote.Enums;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Windows.Forms;
using System.Xml;

using OneNote = Microsoft.Office.Interop.OneNote;

namespace Evernote2Onenote
{
    /// <summary>
    /// main dialog
    /// </summary>
    public partial class MainFrm : Form
    {
        private string _evernoteNotebookPath;
        private readonly SynchronizationContext _synchronizationContext;
        private bool _cancelled;
        private SyncStep _syncStep = SyncStep.Start;
        private Microsoft.Office.Interop.OneNote.Application _onApp;
        private readonly string _xmlNewOutlineContent =
            "<one:Meta name=\"{2}\" content=\"{1}\"/>" +
            "<one:OEChildren><one:HTMLBlock><one:Data><![CDATA[{0}]]></one:Data></one:HTMLBlock>{3}</one:OEChildren>";

        private const string XmlSourceUrl = "<one:OE alignment=\"left\" quickStyleIndex=\"2\"><one:T><![CDATA[From &lt;<a href=\"{0}\">{0}</a>&gt; ]]></one:T></one:OE>";

        private const string XmlNewOutline = "<?xml version=\"1.0\"?>" + "<one:Page xmlns:one=\"{2}\" ID=\"{1}\" dateTime=\"{5}\">" + "<one:Title selected=\"partial\" lang=\"en-US\">" + "<one:OE creationTime=\"{5}\" lastModifiedTime=\"{5}\">" + "<one:T><![CDATA[{3}]]></one:T> " + "</one:OE>" + "</one:Title>{4}" + "<one:Outline>{0}</one:Outline></one:Page>";

        private const string Xmlns = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        private string _enNotebookName = "";
        private bool _useUnfiledSection;
        private string _enexfile = "";
        string _newnbId = "";

        private readonly string _cmdNoteBook = "";
        private DateTime _cmdDate = new DateTime(0);

        private readonly Regex _rxStyle = new Regex("(?<text>\\<(?:div|span).)style=\\\"[^\\\"]*\\\"", RegexOptions.IgnoreCase);
        private readonly Regex _rxCdata = new Regex(@"<!\[CDATA\[<\?xml version=[""']1.0[""'][^?]*\?>", RegexOptions.IgnoreCase);
        private readonly Regex _rxCdata2 = new Regex(@"<!\[CDATA\[<!DOCTYPE en-note \w+ ""https?://xml.evernote.com/pub/enml2.dtd"">", RegexOptions.IgnoreCase);
        private readonly Regex _rxCdataInner = new Regex(@"\<\!\[CDATA\[(?<text>.*)\]\]\>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private readonly Regex _rxBodyStart = new Regex(@"<en-note[^>/]*>", RegexOptions.IgnoreCase);
        private readonly Regex _rxBodyEnd = new Regex(@"</en-note\s*>\s*]]>", RegexOptions.IgnoreCase);
        private readonly Regex _rxBodyEmpty = new Regex(@"<en-note[^>/]*/>\s*]]>", RegexOptions.IgnoreCase);
        private readonly Regex _rxDate = new Regex(@"^date:(.*)$", RegexOptions.IgnoreCase | RegexOptions.Multiline);
        private readonly Regex _rxNote = new Regex("<title>(.+)</title>", RegexOptions.IgnoreCase);
        private readonly Regex _rxComment = new Regex("<!--(.+)-->", RegexOptions.IgnoreCase);
        private static readonly Regex RxDtd = new Regex(@"<!DOCTYPE en-note SYSTEM \""http:\/\/xml\.evernote\.com\/pub\/enml\d*\.dtd\"">", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // This is the constructor for the MainFrm class, which is the main dialog of the application.
        // It takes two parameters: cmdNotebook and cmdDate, which are used to start the synchronization process.
        public MainFrm(string cmdNotebook, string cmdDate)
        {
            // Initialize the form components and get the current synchronization context.
            InitializeComponent();
            _synchronizationContext = SynchronizationContext.Current;

            // Get the version of the application and display it in the versionLabel.
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            versionLabel.Text = $@"Version: {version}";

            // If cmdNotebook is not empty, set _cmdNoteBook to its value.
            if (cmdNotebook.Length > 0)
                _cmdNoteBook = cmdNotebook;

            // If cmdDate is not empty, try to parse it as a DateTime and set _cmdDate to its value.
            if (cmdDate.Length > 0)
            {
                try
                {
                    _cmdDate = DateTime.Parse(cmdDate);
                }
                catch (Exception)
                {
                    MessageBox.Show($"The Datestring\n{cmdDate}\nis not valid!");
                }
            }

            // Set the value of the importDatePicker to _cmdDate, or to its minimum value if it is not a valid date.
            try
            {
                importDatePicker.Value = _cmdDate;
            }
            catch (Exception)
            {
                importDatePicker.Value = importDatePicker.MinDate;
            }

            // If cmdNotebook is not empty, start the synchronization process.
            if (cmdNotebook.Length > 0)
            {
                StartSync();
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }



        // This method updates the progress bar and information text boxes on the main form.
        // It takes two strings as input, line1 and line2, which are used to update the information text boxes.
        // It also takes two integers as input, pos and max, which are used to update the progress bar.
        private void SetInfo(string line1, string line2, int pos, int max)
        {
            var fullpos = 0;

            // The switch statement below calculates the full progress of the synchronization process based on the current step.
            switch (_syncStep)
            {
                // full progress is from 0 - 100'000
                case SyncStep.ExtractNotes:      // 0- 10%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.1) : 0;
                    break;
                case SyncStep.ParseNotes:        // 10-20%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.1) + 10000 : 10000;
                    break;
                case SyncStep.CalculateWhatToDo: // 30-35%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.05) + 30000 : 30000;
                    break;
                case SyncStep.ImportNotes:       // 35-100%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.65) + 35000 : 35000;
                    break;
            }

            // The code below updates the information text boxes and progress bar on the main form.
            _synchronizationContext.Send(delegate
            {
                if (line1 != null)
                    infoText1.Text = line1;
                if (line2 != null)
                    infoText2.Text = line2;
                progressIndicator.Minimum = 0;
                progressIndicator.Maximum = 100000;
                progressIndicator.Value = fullpos;
            }, null);

            // If max is 0, the current step is complete and the next step is started.
            if (max == 0)
                _syncStep++;
        }


        // This method is called when the user clicks the "Import ENEX File" button on the UI.
        private void btnENEXImport_Click(object sender, EventArgs e)
        {
            // If the button text is "Cancel", set the _cancelled flag to true and return.
            if (btnENEXImport.Text == "Cancel")
            {
                _cancelled = true;
                return;
            }

            // Create a new OpenFileDialog object and set its properties.
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = @"Evernote exports|*.enex";
            openFileDialog1.Title = @"Select the ENEX file";
            openFileDialog1.CheckPathExists = true;

            // Show the Dialog and check if the user selected a file.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Set the _enexfile field to the selected file path and start the synchronization process.
                _enexfile = openFileDialog1.FileName;
                StartSync();
            }
        }


        // This method is called when the user clicks the "Import ENEX File" button on the UI.
        // It creates a new notebook in OneNote and starts the synchronization process by parsing the notes from the Evernote export file and importing them to OneNote.
        // If the user has specified a specific notebook and date to synchronize, the method closes the application after the synchronization is complete.
        private void StartSync()
        {
            // If an Evernote export file has been selected, set the notebook name to the name of the export file.
            if (!string.IsNullOrEmpty(_enexfile))
            {
                _enNotebookName = Path.GetFileNameWithoutExtension(_enexfile);
            }

            // If a specific notebook has been specified by the user, set the notebook name to the specified name.
            if (_cmdNoteBook.Length > 0)
                _enNotebookName = _cmdNoteBook;

            try
            {
                // Create a new instance of the OneNote application.
                _onApp = new OneNote.Application();
            }
            catch (Exception ex)
            {
                // If an exception occurs while creating the OneNote application instance, display an error message and return.
                MessageBox.Show(
                    $"Could not connect to Onenote!\nReasons for this might be:\n* The desktop version of onenote is not installed\n* Onenote is not installed properly\n* Onenote is already running but with a different user account\n\n{ex}");
                return;
            }

            // If the OneNote application instance is null, display an error message and return.
            if (_onApp == null)
            {
                MessageBox.Show("Could not connect to Onenote!\nReasons for this might be:\n* The desktop version of onenote is not installed\n* Onenote is not installed properly\n* Onenote is already running but with a different user account\n");
                return;
            }

            // Create a new notebook in OneNote with the specified name.
            try
            {
                // Get the hierarchy for the default notebook folder.
                _onApp.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out var xmlHierarchy);

                // Get the path to the default notebook folder and append the notebook name to it.
                _onApp.GetSpecialLocation(OneNote.SpecialLocation.slDefaultNotebookFolder, out _evernoteNotebookPath);
                _evernoteNotebookPath += "\\" + _enNotebookName;

                // Open the new notebook and get the hierarchy for its pages.
                _onApp.OpenHierarchy(_evernoteNotebookPath, "", out var newnbId, OneNote.CreateFileType.cftNotebook);
                _onApp.GetHierarchy(newnbId, OneNote.HierarchyScope.hsPages, out _);

                // Load and process the hierarchy.
                var docHierarchy = new XmlDocument();
                docHierarchy.LoadXml(xmlHierarchy);
                var hierarchy = new StringBuilder();
                AppendHierarchy(docHierarchy.DocumentElement, hierarchy, 0);
            }
            catch (Exception)
            {
                // If an exception occurs while creating the new notebook, try to create it in the Unfiled Notes section instead.
                try
                {
                    // Get the hierarchy for the Unfiled Notes section.
                    _onApp.GetHierarchy("", OneNote.HierarchyScope.hsPages, out _);

                    // Get the path to the Unfiled Notes section and open it.
                    _onApp.GetSpecialLocation(OneNote.SpecialLocation.slUnfiledNotesSection, out _evernoteNotebookPath);
                    _onApp.OpenHierarchy(_evernoteNotebookPath, "", out _newnbId);
                    _onApp.GetHierarchy(_newnbId, OneNote.HierarchyScope.hsPages, out _);

                    // Set the flag to indicate that the Unfiled Notes section is being used.
                    _useUnfiledSection = true;
                }
                catch (Exception ex2)
                {
                    // If an exception occurs while creating the new notebook in the Unfiled Notes section, display an error message and return.
                    MessageBox.Show($"Could not create the target notebook in Onenote!\n{ex2}");
                    return;
                }
            }

            // If the date specified by the user is later than the current date, set the current date to the specified date.
            if (importDatePicker.Value > _cmdDate)
                _cmdDate = importDatePicker.Value;

            // If the "Import ENEX File" button is clicked, start the synchronization process.
            if (btnENEXImport.Text == "Import ENEX File")
            {
                // Change the text of the "Import ENEX File" button to "Cancel".
                btnENEXImport.Text = "Cancel";

                // Create a delegate for the ImportNotesToOnenote method and begin invoking it asynchronously.
                MethodInvoker syncDelegate = ImportNotesToOnenote;
                syncDelegate.BeginInvoke(null, null);
            }
            else
            {
                // If the "Cancel" button is clicked, set the cancelled flag to true.
                _cancelled = true;
            }
        }


        // This method is called when the user clicks the "Import ENEX File" button on the UI.
        // It starts the synchronization process by parsing the notes from the Evernote export file and importing them to OneNote.
        // If the user has specified a specific notebook and date to synchronize, the method closes the application after the synchronization is complete.
        private void ImportNotesToOnenote()
        {
            _syncStep = SyncStep.Start;

            // If an Evernote export file has been selected, parse the notes from it.
            if (!string.IsNullOrEmpty(_enexfile))
            {
                var notesEvernote = new List<Note>();
                if (_enexfile != string.Empty)
                {
                    SetInfo("Parsing notes from Evernote", "", 0, 0);
                    notesEvernote = ParseNotes(_enexfile);
                }

                // Import the parsed notes to OneNote.
                if (_enexfile != string.Empty)
                {
                    SetInfo("importing notes to Onenote", "", 0, 0);
                    ImportNotesToOnenote(notesEvernote, _enexfile);
                }
            }

            // Reset the _enexfile variable and display the appropriate message in the UI.
            _enexfile = "";
            if (_cancelled)
            {
                SetInfo(null, "Operation cancelled", 0, 0);
            }
            else
                SetInfo("", "", 0, 0);

            // Update the UI to indicate that the synchronization is complete.
            _synchronizationContext.Send(delegate
            {
                btnENEXImport.Text = @"Import ENEX File";
                infoText1.Text = @"Finished";
                progressIndicator.Minimum = 0;
                progressIndicator.Maximum = 100000;
                progressIndicator.Value = 0;
            }, null);

            // If the user has specified a specific notebook and date to synchronize, close the application.
            if (_cmdNoteBook.Length > 0)
            {
                _synchronizationContext.Send(delegate
                {
                    Close();
                }, null);
            }
        }



        // This method parses the notes from an Evernote export file and returns a list of Note objects.
        // It takes a string parameter, exportFile, which is the path to the Evernote export file.
        // It uses an XmlTextReader to read the XML data from the export file and creates a new Note object for each <note> element.
        // The method returns a List<Note> object containing all the Note objects parsed from the export file.
        private List<Note> ParseNotes(string exportFile)
        {
            _syncStep = SyncStep.ParseNotes;
            var noteList = new List<Note>();
            if (_cancelled)
            {
                return noteList;
            }

            // Create a new XmlTextReader to read the XML data from the export file
            var xtrInput = new XmlTextReader(exportFile);
            var xmltext = "";
            try
            {
                // Loop through the XML data and create a new Note object for each <note> element
                while (xtrInput.Read())
                {
                    while ((xtrInput.NodeType == XmlNodeType.Element) && (xtrInput.Name.ToLower() == "note"))
                    {
                        if (_cancelled)
                        {
                            break;
                        }

                        // Load the <note> element into an XmlDocument object
                        var xmlDocItem = new XmlDocument();
                        xmltext = SanitizeXml(xtrInput.ReadOuterXml());
                        xmlDocItem.LoadXml(xmltext);
                        var node = xmlDocItem.FirstChild;

                        // node is <note> element
                        // node.FirstChild.InnerText is <title>
                        node = node.FirstChild;

                        // Create a new Note object and set its Title property to the value of the <title> element
                        var note = new Note
                        {
                            Title = HttpUtility.HtmlDecode(node.InnerText)
                        };

                        // Add the new Note object to the noteList
                        noteList.Add(note);
                    }
                }

                // Close the XmlTextReader
                xtrInput.Close();
            }
            catch (XmlException ex)
            {
                // Handle any XmlExceptions that occur during parsing
                // This can happen if the notebook was empty or does not exist, or if a note isn't properly xml encoded
                // Try to find the name of the note that's causing the problems
                var notename = "";
                if (xmltext.Length > 0)
                {
                    var notematch = _rxNote.Match(xmltext);
                    if (notematch.Groups.Count == 2)
                    {
                        notename = notematch.Groups[1].ToString();
                    }
                }

                // Create a temporary directory to store the failed note
                var temppath = Path.GetTempPath() + "\\ev2on";
                var tempfilepathDir = temppath + "\\failedNotes";
                try
                {
                    Directory.CreateDirectory(tempfilepathDir);
                    var tempfilepath = tempfilepathDir + "\\note-";
                    tempfilepath += Guid.NewGuid().ToString();
                    tempfilepath += ".xml";
                    File.WriteAllText(tempfilepath, xmltext);
                }
                catch (Exception)
                {
                    // ignored
                }

                // Display an error message to the user with information about the failed note and a link to create an issue on GitHub
                MessageBox.Show(notename.Length > 0
                    ? $"Error parsing the note \"{notename}\" in notebook \"{_enNotebookName}\",\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues"
                    : $"Error parsing the notebook \"{_enNotebookName}\"\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues");
            }

            // Return the list of Note objects parsed from the export file
            return noteList;
        }

        private void ImportNotesToOnenote(List<Note> notesEvernote, string exportFile)
        {
            _syncStep = SyncStep.CalculateWhatToDo;
            var uploadcount = notesEvernote.Count;

            var temppath = Path.GetTempPath() + "\\ev2on";
            Directory.CreateDirectory(temppath);

            _syncStep = SyncStep.ImportNotes;
            var counter = 0;


            {
                var xmltext = "";
                try
                {
                    var xtrInput = new XmlTextReader(exportFile);
                    while (xtrInput.Read())
                    {
                        while ((xtrInput.NodeType == XmlNodeType.Element) && (xtrInput.Name.ToLower() == "note"))
                        {
                            if (_cancelled)
                            {
                                break;
                            }

                            var xmlDocItem = new XmlDocument();
                            xmltext = SanitizeXml(xtrInput.ReadOuterXml());
                            xmlDocItem.LoadXml(xmltext);
                            var node = xmlDocItem.FirstChild;

                            // node is <note> element
                            // This code parses an Evernote XML file and extracts the relevant information to create a Note object.
                            // The Note object is then used to create a new note in OneNote.

                            // node.FirstChild.InnerText is <title>
                            node = node.FirstChild;

                            // Create a new Note object
                            var note = new Note
                            {
                                // Set the title of the note to the decoded value of the <title> element
                                Title = HttpUtility.HtmlDecode(node.InnerText)
                            };

                            // If the title starts with "=?", it is encoded using RFC 2047 and needs to be decoded
                            if (note.Title.StartsWith("=?"))
                                note.Title = Rfc2047Decoder.Parse(note.Title);

                            // Get the <content> element
                            var contentElements = xmlDocItem.GetElementsByTagName("content");
                            if (contentElements.Count > 0)
                            {
                                node = contentElements[0];
                            }

                            // Set the content of the note to the decoded value of the <content> element
                            note.Content = HttpUtility.HtmlDecode(node.InnerXml);

                            // If the content starts with "=?", it is encoded using RFC 2047 and needs to be decoded
                            if (note.Content.StartsWith("=?"))
                                note.Content = Rfc2047Decoder.Parse(note.Content);

                            // Get all <resource> elements (attachments)
                            var atts = xmlDocItem.GetElementsByTagName("resource");
                            foreach (XmlNode xmln in atts)
                            {
                                // Create a new Attachment object
                                var attachment = new Attachment
                                {
                                    // Set the Base64Data property to the value of the <data> element
                                    Base64Data = xmln.FirstChild.InnerText
                                };

                                // Compute the MD5 hash of the attachment data
                                var data = Convert.FromBase64String(xmln.FirstChild.InnerText);
                                var hash = new System.Security.Cryptography.MD5CryptoServiceProvider().ComputeHash(data);
                                var hashHex = BitConverter.ToString(hash).Replace("-", string.Empty).ToLower();

                                // Set the Hash property to the computed hash
                                attachment.Hash = hashHex;

                                // Get the <file-name> element for this attachment
                                var fns = xmlDocItem.GetElementsByTagName("file-name");
                                if (fns.Count > note.Attachments.Count)
                                {
                                    // Set the FileName property to the decoded value of the <file-name> element
                                    attachment.FileName = HttpUtility.HtmlDecode(fns.Item(note.Attachments.Count).InnerText);

                                    // If the file name starts with "=?", it is encoded using RFC 2047 and needs to be decoded
                                    if (attachment.FileName.StartsWith("=?"))
                                        attachment.FileName = Rfc2047Decoder.Parse(attachment.FileName);

                                    // Remove any invalid characters from the file name
                                    var invalid = new string(Path.GetInvalidFileNameChars());
                                    foreach (var c in invalid)
                                    {
                                        attachment.FileName = attachment.FileName.Replace(c.ToString(), "");
                                    }

                                    // Escape any special characters in the file name
                                    attachment.FileName = System.Security.SecurityElement.Escape(attachment.FileName);
                                }

                                // Get the <mime> element for this attachment
                                var mimes = xmlDocItem.GetElementsByTagName("mime");
                                if (mimes.Count > note.Attachments.Count)
                                {
                                    // Set the ContentType property to the decoded value of the <mime> element
                                    attachment.ContentType = HttpUtility.HtmlDecode(mimes.Item(note.Attachments.Count).InnerText);
                                }

                                // Add the attachment to the note's Attachments collection
                                note.Attachments.Add(attachment);
                            }

                            // Get all <tag> elements
                            var tagslist = xmlDocItem.GetElementsByTagName("tag");
                            foreach (XmlNode n in tagslist)
                            {
                                // Add the decoded value of the <tag> element to the note's Tags collection
                                note.Tags.Add(HttpUtility.HtmlDecode(n.InnerText));
                            }

                            // Get the <created> element
                            var datelist = xmlDocItem.GetElementsByTagName("created");
                            foreach (XmlNode n in datelist)
                            {
                                // If the <created> element is in the correct format, set the Date property of the note to its value
                                if (DateTime.TryParseExact(n.InnerText, "yyyyMMddTHHmmssZ", CultureInfo.CurrentCulture, DateTimeStyles.AdjustToUniversal, out var dateCreated))
                                {
                                    note.Date = dateCreated;
                                }
                            }

                            // If the "modifiedDateCheckbox" is checked, get the <updated> element
                            if (modifiedDateCheckbox.Checked)
                            {
                                var datelist2 = xmlDocItem.GetElementsByTagName("updated");
                                foreach (XmlNode n in datelist2)
                                {
                                    // If the <updated> element is in the correct format, set the Date property of the note to its value
                                    if (DateTime.TryParseExact(n.InnerText, "yyyyMMddTHHmmssZ", CultureInfo.CurrentCulture, DateTimeStyles.AdjustToUniversal, out var dateUpdated))
                                    {
                                        note.Date = dateUpdated;
                                    }
                                }
                            }

                            // Get the <source-url> element
                            var sourceurl = xmlDocItem.GetElementsByTagName("source-url");
                            note.SourceUrl = "";
                            foreach (XmlNode n in sourceurl)
                            {
                                try
                                {
                                    // If the <source-url> element starts with "file://" or "en-cache://", skip it
                                    if (n.InnerText.StartsWith("file://"))
                                        continue;
                                    if (n.InnerText.StartsWith("en-cache://"))
                                        continue;

                                    // Set the SourceUrl property of the note to the value of the <source-url> element
                                    note.SourceUrl = n.InnerText;
                                }
                                catch (FormatException)
                                {
                                    // If the <source-url> element is not in the correct format, ignore it
                                }
                            }

                            if (_cmdDate > note.Date)
                                continue;

                            SetInfo(null, $"importing note ({counter + 1} of {uploadcount}) : \"{note.Title}\"", counter++, uploadcount);

                            var htmlBody = note.Content;

                            var tempfiles = new List<string>();
                            var xmlAttachments = "";
                            foreach (var attachment in note.Attachments)
                            {
                                // save the attached file
                                var tempfilepath = temppath + "\\";
                                var data = Convert.FromBase64String(attachment.Base64Data);
                                tempfilepath += attachment.Hash;
                                Stream fs = new FileStream(tempfilepath, FileMode.Create);
                                fs.Write(data, 0, data.Length);
                                fs.Close();
                                tempfiles.Add(tempfilepath);

                                var rx = new Regex(@"<en-media\b[^>]*?hash=""" + attachment.Hash + @"""[^>]*/>", RegexOptions.IgnoreCase);
                                if ((attachment.ContentType != null) && (attachment.ContentType.Contains("image") && rx.Match(htmlBody).Success))
                                {
                                    // replace the <en-media /> tag with an <img /> tag
                                    htmlBody = rx.Replace(htmlBody, @"<img src=""file:///" + tempfilepath + @"""/>");
                                }
                                else
                                {
                                    rx = new Regex(@"<en-media\b[^>]*?hash=""" + attachment.Hash + @"""[^>]*></en-media>", RegexOptions.IgnoreCase);
                                    if ((attachment.ContentType != null) && (attachment.ContentType.Contains("image") && rx.Match(htmlBody).Success))
                                    {
                                        // replace the <en-media /> tag with an <img /> tag
                                        htmlBody = rx.Replace(htmlBody, @"<img src=""file:///" + tempfilepath + @"""/>");
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(attachment.FileName))
                                        {
                                            // do not attach proxy.php image files: those are overlay images created by evernote to search text in images
                                            if (!attachment.ContentType.Contains("image") || attachment.FileName != "proxy.php")
                                                xmlAttachments +=
                                                    $"<one:InsertedFile pathSource=\"{tempfilepath}\" preferredName=\"{attachment.FileName}\" />";
                                        }
                                        else
                                            xmlAttachments +=
                                                $"<one:InsertedFile pathSource=\"{tempfilepath}\" preferredName=\"{attachment.Hash}\" />";
                                    }
                                }
                            }
                            note.Attachments.Clear();

                            htmlBody = _rxStyle.Replace(htmlBody, "${text}");
                            htmlBody = _rxComment.Replace(htmlBody, string.Empty);
                            htmlBody = _rxCdata.Replace(htmlBody, string.Empty);
                            htmlBody = _rxCdata2.Replace(htmlBody, string.Empty);
                            htmlBody = RxDtd.Replace(htmlBody, string.Empty);
                            htmlBody = _rxBodyStart.Replace(htmlBody, "<body>");
                            htmlBody = _rxBodyEnd.Replace(htmlBody, "</body>");
                            htmlBody = _rxBodyEmpty.Replace(htmlBody, "<body></body>");
                            htmlBody = htmlBody.Trim();
                            htmlBody = @"<!DOCTYPE html><head></head>" + htmlBody;

                            var emailBody = htmlBody;
                            emailBody = _rxDate.Replace(emailBody, "Date: " + note.Date.ToString("ddd, dd MMM yyyy HH:mm:ss K"));
                            emailBody = emailBody.Replace("&apos;", "'");
                            emailBody = emailBody.Replace("’", "'");
                            emailBody = _rxCdataInner.Replace(emailBody, "&lt;![CDATA[${text}]]&gt;");
                            emailBody = emailBody.Replace("‘", "'");

                            try
                            {
                                var pageId = string.Empty;

                                // Place the note in one section: use the first tag if available, otherwise "not specified".
                                // All tags are preserved as a "Tags:" line in the note body.
                                var sectionName = (note.Tags.Count > 0 && !_useUnfiledSection)
                                    ? note.Tags[0]
                                    : "not specified";
                                var sectionId = _useUnfiledSection ? _newnbId : GetSection(sectionName);
                                _onApp.CreateNewPage(sectionId, out pageId, OneNote.NewPageStyle.npsBlankPageWithTitle);
                                var outlineId = new Random().Next();
                                var xmlSource = string.Format(XmlSourceUrl, note.SourceUrl);
                                var xmlTagsLine = note.Tags.Count > 0
                                    ? string.Format("<one:OE alignment=\"left\" quickStyleIndex=\"2\"><one:T><![CDATA[Tags: {0}]]></one:T></one:OE>",
                                        string.Join(", ", note.Tags))
                                    : "";
                                var extraContent = (note.SourceUrl.Length > 0 ? xmlSource : "") + xmlTagsLine;
                                var outlineContent = string.Format(_xmlNewOutlineContent, emailBody, outlineId, System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"), extraContent);
                                var xml = string.Format(XmlNewOutline, outlineContent, pageId, Xmlns, System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"), xmlAttachments, note.Date.ToString("yyyy'-'MM'-'ddTHH':'mm':'ss'Z'"));
                                _onApp.UpdatePageContent(xml, DateTime.MinValue, OneNote.XMLSchema.xs2013, true);
                                _onApp.SyncHierarchy(pageId);
                            }
                            catch (Exception ex)
                            {
                                var tempfilepathDir = temppath + "\\failedNotes";
                                try
                                {
                                    Directory.CreateDirectory(tempfilepathDir);
                                    var tempfilepath = tempfilepathDir + "\\note-";
                                    tempfilepath += Guid.NewGuid().ToString();
                                    tempfilepath += ".xml";
                                    File.WriteAllText(tempfilepath, xmltext);
                                }
                                catch (Exception)
                                {
                                    // ignored
                                }

                                MessageBox.Show($"Note:{note.Title}\n{ex}\n\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues");
                            }

                            foreach (var p in tempfiles)
                            {
                                File.Delete(p);
                            }
                        }
                    }

                    xtrInput.Close();
                }
                catch (XmlException ex)
                {
                    // happens if the notebook was empty or does not exist.
                    // Or due to a parsing error if a note isn't properly xml encoded
                    // try to find the name of the note that's causing the problems
                    var notename = "";
                    if (xmltext.Length > 0)
                    {
                        var notematch = _rxNote.Match(xmltext);
                        if (notematch.Groups.Count == 2)
                        {
                            notename = notematch.Groups[1].ToString();
                        }
                    }
                    var tempfilepathDir = temppath + "\\failedNotes";
                    try
                    {
                        Directory.CreateDirectory(tempfilepathDir);
                        var tempfilepath = tempfilepathDir + "\\note-";
                        tempfilepath += Guid.NewGuid().ToString();
                        tempfilepath += ".xml";
                        File.WriteAllText(tempfilepath, xmltext);
                    }
                    catch (Exception)
                    {
                        // ignored
                    }

                    MessageBox.Show(notename.Length > 0
                        ? $"Error parsing the note \"{notename}\" in notebook \"{_enNotebookName}\",\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues"
                        : $"Error parsing the notebook \"{_enNotebookName}\"\n{ex}\\n\\nA copy of the note is left in {tempfilepathDir}. If you want to help fix the problem, please consider creating an issue and attaching that note to it: https://github.com/stefankueng/EvImSync/issues");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Exception importing notes:\n{ex}");
                }
            }
            _onApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        

        // This method appends the hierarchy of OneNote elements to a StringBuilder object.
        // The hierarchy includes elements such as Notebooks, SectionGroups, Sections, and Pages.
        // The hierarchy is represented as a string with each element on a new line, indented according to its level in the hierarchy.
        // The string is used to display the hierarchy in the UI and to help the user select the destination for imported notes.
        private void AppendHierarchy(XmlNode xml, StringBuilder str, int level)
        {
            // The set of elements that are themselves meaningful to export:
            if (xml.Name == "one:Notebook" || xml.Name == "one:SectionGroup" || xml.Name == "one:Section" || xml.Name == "one:Page")
            {
                // If the element has an ID and a name attribute, append its information to the StringBuilder.
                if (xml.Attributes != null)
                {
                    // If the element is a Section and its path attribute matches the Evernote notebook path, set its ID to "UnfiledNotes".
                    var id = xml.Attributes != null && xml.LocalName == "Section" && xml.Attributes["path"].Value == _evernoteNotebookPath
                        ? "UnfiledNotes"
                        : xml.Attributes["ID"].Value;
                    // Encode the name attribute to prevent HTML injection attacks.
                    var name = HttpUtility.HtmlEncode(xml.Attributes["name"].Value);
                    // Append the element's information to the StringBuilder, with indentation based on its level in the hierarchy.
                    if (str.Length > 0)
                        str.Append("\n");
                    str.Append($"{level.ToString()} {xml.LocalName} {id} {name}");
                }
            }
            // The set of elements that contain children that are meaningful to export:
            if (xml.Name == "one:Notebooks" || xml.Name == "one:Notebook" || xml.Name == "one:SectionGroup" || xml.Name == "one:Section")
            {
                // Recursively call this method on each child element.
                foreach (XmlNode child in xml.ChildNodes)
                {
                    int nextLevel;
                    if (xml.Name == "one:Notebooks")
                        nextLevel = level;
                    else
                        nextLevel = level + 1;
                    AppendHierarchy(child, str, nextLevel);
                }
            }
        }


        private void homeLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://tools.stefankueng.com/Evernote2Onenote.html");
        }
        private string GetSection(string sectionName)
        {
            var newnbId = "";
            try
            {
                // remove and/or replace characters that are not allowed in Onenote section names
                sectionName = sectionName.Replace("?", "");
                sectionName = sectionName.Replace("*", "");
                sectionName = sectionName.Replace("/", "");
                sectionName = sectionName.Replace("\\", "");
                sectionName = sectionName.Replace(":", "");
                sectionName = sectionName.Replace("<", "");
                sectionName = sectionName.Replace(">", "");
                sectionName = sectionName.Replace("|", "");
                sectionName = sectionName.Replace("&", "");
                sectionName = sectionName.Replace("#", "");
                sectionName = sectionName.Replace("\"", "'");
                sectionName = sectionName.Replace("%", "");

                _onApp.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out var xmlHierarchy);

                _onApp.OpenHierarchy(_evernoteNotebookPath + "\\" + sectionName + ".one", "", out newnbId, OneNote.CreateFileType.cftSection);
                _onApp.GetHierarchy(newnbId, OneNote.HierarchyScope.hsSections, out _);

                // Load and process the hierarchy
                var docHierarchy = new XmlDocument();
                docHierarchy.LoadXml(xmlHierarchy);
                var hierarchy = new StringBuilder(sectionName);
                AppendHierarchy(docHierarchy.DocumentElement, hierarchy, 0);
            }
            catch (Exception /*ex*/)
            {
                //MessageBox.Show(string.Format("Exception creating section \"{0}\":\n{1}", sectionName, ex.ToString()));
            }
            return newnbId;
        }

        private string SanitizeXml(string text)
        {
            //text = HttpUtility.HtmlDecode(text);
            var rxtitle = new Regex("<note><title>(.+)</title>", RegexOptions.IgnoreCase);
            var match = rxtitle.Match(text);
            if (match.Groups.Count == 2)
            {
                var title = match.Groups[1].ToString();
                title = title.Replace("&", "&amp;");
                title = title.Replace("\"", "&quot;");
                title = title.Replace("'", "&apos;");
                title = title.Replace("’", "&apos;");
                title = title.Replace("<", "&lt;");
                title = title.Replace(">", "&gt;");
                title = title.Replace("@", "&#64;");
                text = rxtitle.Replace(text, "<note><title>" + title + "</title>");
            }

            var rxauthor = new Regex("<author>(.+)</author>", RegexOptions.IgnoreCase);
            var authormatch = rxauthor.Match(text);
            if (match.Groups.Count == 2)
            {
                var author = authormatch.Groups[1].ToString();
                author = author.Replace("&", "&amp;");
                author = author.Replace("\"", "&quot;");
                author = author.Replace("'", "&apos;");
                author = author.Replace("’", "&apos;");
                author = author.Replace("<", "&lt;");
                author = author.Replace(">", "&gt;");
                author = author.Replace("@", "&#64;");
                text = rxauthor.Replace(text, "<author>" + author + "</author>");
            }

            var rxfilename = new Regex("<file-name>(.+)</file-name>", RegexOptions.IgnoreCase);
            if (match.Groups.Count == 2)
            {
                MatchEvaluator myEvaluator = FilenameMatchEvaluator;
                text = rxfilename.Replace(text, myEvaluator);
            }

            return text;
        }

        // This method replaces invalid characters in a filename with an empty string and escapes any special characters.
        // It is used as a MatchEvaluator for the rxfilename regex in the SanitizeXml method.
        private string FilenameMatchEvaluator(Match m)
        {
            var filename = m.Groups[1].ToString();
            filename = filename.Replace("&nbsp;", " ");
            // remove illegal path chars
            var invalid = new string(Path.GetInvalidFileNameChars());
            foreach (var c in invalid)
            {
                filename = filename.Replace(c.ToString(), "");
            }
            filename = System.Security.SecurityElement.Escape(filename);
            return "<file-name>" + filename + "</file-name>";
        }

        // This event handler is called when the MainFrm form is closing.
        // It sets the _cancelled flag to true and performs garbage collection.
        private void MainFrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_onApp != null)
            {
                _cancelled = true;
            }
            _onApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
