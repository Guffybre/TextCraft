using System;
using System.ClientModel;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;
using OpenAI.Images;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.Exceptions;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class RAGControl : UserControl
    {
        // Public
        public static readonly int CHUNK_LEN = CommonUtils.TokensToCharCount(256);

        // Private
        private static readonly CultureLocalizationHelper _cultureHelper = new CultureLocalizationHelper("TextForge.RAGControl", typeof(RAGControl).Assembly);
        private ToolTip _fileToolTip = new ToolTip();
        private Queue<string> _removalQueue = new Queue<string>();
        private ConcurrentDictionary<int, int> _indexFileCount = new ConcurrentDictionary<int, int>();
        private BindingList<KeyValuePair<string, string>> _fileList; // Use KeyValuePair for label and filename
        private HyperVectorDB.HyperVectorDB _db;
        private bool _isIndexing;
        private float preciseProgressBar = 0;
        private readonly object progressBarLock = new object();

        public RAGControl()
        {
            try
            {
                InitializeComponent();
                this.Load += (s, e) =>
                {
                    // Run the background task to initialize BindingList and FileListBox
                    Task.Run(() => InitializeRAGControl());
                };
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void InitializeRAGControl()
        {
            FileListBox.Invoke(new Action(() =>
            {
                _fileList = new BindingList<KeyValuePair<string, string>>();
                FileListBox.DataSource = _fileList;
                FileListBox.DisplayMember = "Key";  // Display the label (Key)
                FileListBox.ValueMember = "Value";  // Internally use the filename (Value)

                _fileToolTip.ShowAlways = true;     // Always show the tooltip

                // Attach MouseMove event to FileListBox to display the full path in the tooltip
                FileListBox.MouseMove += FileListBox_MouseMove;
            }));

            lock (Forge.InitializeDoor)
            {
                _db = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
            }
        }

        private async void AddButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog() { Title = _cultureHelper.GetLocalizedString("[AddButton_Click] OpenFileDialog #1 Title"), Filter = "PDF files (*.pdf)|*.pdf", Multiselect = true })
                {
                    List<string> filesToIndex = new List<string>();
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        foreach (string fileName in openFileDialog.FileNames)
                        {
                            if (!_fileList.Any(file => file.Value == fileName))
                            {
                                _fileList.Add(new KeyValuePair<string, string>(@$"📄 {Path.GetFileName(fileName)}", fileName));
                                filesToIndex.Add(fileName);
                                if (!RemoveButton.Enabled)
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        RemoveButton.Enabled = true;
                                    });
                                }
                            }
                        }

                        ChangeProgressBarVisibility(true);
                        
                        int dictCount = _indexFileCount.Count;
                        _indexFileCount.TryAdd(dictCount, filesToIndex.Count);
                        lock (progressBarLock)
                        {
                            SetProgressBarValue(GetProgressBarValue() / (dictCount + 1));
                        }

                        foreach (var filePath in filesToIndex)
                            await IndexDocumentAsync(filePath);

                        int temp;
                        _indexFileCount.TryRemove(dictCount, out temp);

                        if (_indexFileCount.Count == 0)
                            await ChangeProgressBarVisibilityAfterSleep(2, false);
                    }
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void RemoveButton_Click(object sender, EventArgs e)
        {
            try
            {
                RemoveSelectedDocument();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void FileListBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Delete && FileListBox.Items.Count > 0)
                    RemoveSelectedDocument();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void FileListBox_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                // Get the index of the item under the mouse cursor
                int index = FileListBox.IndexFromPoint(e.Location);
                if (index != ListBox.NoMatches)
                {
                    // Get the KeyValuePair (label, file path) for the item
                    var item = (KeyValuePair<string, string>)FileListBox.Items[index];

                    // Show the file path in the tooltip
                    _fileToolTip.SetToolTip(FileListBox, item.Value);
                }
                else
                {
                    // Clear the tooltip if not hovering over an item
                    _fileToolTip.SetToolTip(FileListBox, string.Empty);
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void RemoveSelectedDocument()
        {
            string selectedDocument = FileListBox.SelectedItem.ToString();
            if (_isIndexing)
            {
                if (!_removalQueue.Contains(selectedDocument))
                    _removalQueue.Enqueue(selectedDocument);
            }
            else
            {
                DeleteDocument(selectedDocument);
            }
            _fileList.RemoveAt(FileListBox.SelectedIndex);
            AutoHideRemoveButton();
        }

        private async Task IndexDocumentAsync(string filePath)
        {
            IEnumerable<string> fileContent;
            try
            {
                fileContent = await ReadPdfFileAsync(filePath, CHUNK_LEN);
            }
            catch
            {
                this.Invoke((MethodInvoker)delegate
                {
                    // Find and remove the file entry from _fileList based on the internal filename (filePath)
                    var fileEntry = _fileList.FirstOrDefault(file => file.Value == filePath);
                    if (fileEntry.Key != null) // If the file entry is found
                    {
                        _fileList.Remove(fileEntry);
                    }

                    // Automatically hide the remove button if there are no more files in the list
                    AutoHideRemoveButton();
                });
                throw;
            }

            _db.CreateIndex(filePath);
            await Task.Run(() => {
                _isIndexing = true;

                int fileContentCount = fileContent.Count();
                int progressBarIncrement = (int)(fileContentCount * 0.1);
                for (int i = 0; i < fileContentCount; i++)
                {
                    AddDocument(filePath, fileContent.ElementAt(i));
                    if (i % progressBarIncrement == 0)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            UpdateProgressBar((float)progressBarIncrement / fileContentCount);
                        });
                    }
                }

                _isIndexing = false;

                // Process any queued removal requests
                ProcessRemovalQueue();
            });
        }

        private bool AddDocument(string filePath, string content)
        {
            return _db.IndexDocument(filePath, content);
        }

        private bool DeleteDocument(string filePath)
        {
            return _db.DeleteIndex(filePath);
        }

        private void AutoHideRemoveButton()
        {
            if (_fileList.Count == 0)
                RemoveButton.Enabled = false;
        }

        private async Task ChangeProgressBarVisibilityAfterSleep(int seconds, bool val)
        {
            await Task.Delay(seconds * 1000);
            this.Invoke((MethodInvoker)delegate
            {
                ChangeProgressBarVisibility(val);
                ResetProgressBar();
            });
        }

        private void ChangeProgressBarVisibility(bool val)
        {
            this.progressBar1.Visible = val;
        }

        private void ResetProgressBar()
        {
            preciseProgressBar = 0;
            SetProgressBarValue(0);
        }

        private float GetProgressBarValue()
        {
            return this.progressBar1.Value / ((float)this.progressBar1.Maximum);
        }

        private void SetProgressBarValue(float val)
        {
            preciseProgressBar = val;
            this.progressBar1.Value = (int)(val * this.progressBar1.Maximum);
        }

        private void UpdateProgressBar(float val)
        {
            lock (progressBarLock)
            {
                int maxProgress = this.progressBar1.Maximum;
                preciseProgressBar += val / GetIndexFileCount();

                // Clipping
                if (preciseProgressBar > 1)
                    SetProgressBarValue(1);
                else
                    SetProgressBarValue(preciseProgressBar);
            }
        }

        private int GetIndexFileCount()
        {
            int fileCount = 0;
            foreach (var count in _indexFileCount)
                fileCount += count.Value;
            return fileCount;
        }

        private void ProcessRemovalQueue()
        {
            int initialQueueCount = _removalQueue.Count;

            for (int i = 0; i < initialQueueCount; i++)
            {
                string documentToRemove = _removalQueue.Dequeue();

                // Check if the document (by filename) exists in the _fileList
                var fileEntry = _fileList.FirstOrDefault(file => file.Value == documentToRemove);

                // If the document is found, attempt to remove it
                if (fileEntry.Key != null)
                {
                    if (!DeleteDocument(documentToRemove)) // Try removing the document
                    {
                        // If removal fails, re-enqueue the document and stop processing
                        _removalQueue.Enqueue(documentToRemove);
                        break;
                    }
                }
            }
        }

        public static async Task<IEnumerable<string>> ReadPdfFileAsync(string filePath, int chunkLen)
        {
            return await Task.Run(() =>
            {
                List<string> chunks = new List<string>();
                try
                {
                    PdfDocument doc;
                    try { doc = PdfDocument.Open(filePath); }
                    catch (PdfDocumentEncryptedException) { throw new ArgumentException(); }

                    try { IteratePdfFile(ref doc, ref chunks, chunkLen); }
                    finally { doc.Dispose(); }
                }
                catch (ArgumentException)
                {
                    PasswordPrompt passwordDialog = new PasswordPrompt();
                    if (passwordDialog.ShowDialog() == DialogResult.OK)
                    {
                        PdfDocument unlockedDoc = PdfDocument.Open(filePath, new ParsingOptions { Password = passwordDialog.Password });
                        try { IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen); } 
                        finally { unlockedDoc.Dispose(); }
                    }
                    else
                    {
                        throw new InvalidDataException(_cultureHelper.GetLocalizedString("[ReadPdfFileAsync] InvalidDataException #1"));
                    }
                }
                return chunks;
            });
        }

        private static void IteratePdfFile(ref PdfDocument document, ref List<string> chunks, int chunkLen)
        {
            IterateInnerPdfFile(ref document, ref chunks, chunkLen);

            IReadOnlyList<EmbeddedFile> embeddedFiles;
            if (document.Advanced.TryGetEmbeddedFiles(out embeddedFiles))
            {
                foreach (var embeddedFile in embeddedFiles)
                {
                    if (embeddedFile.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            PdfDocument embedDoc;
                            try { embedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray()); }
                            catch (PdfDocumentEncryptedException) { throw new ArgumentException(); }

                            try { IteratePdfFile(ref embedDoc, ref chunks, chunkLen); }
                            finally { embedDoc.Dispose(); }
                        } catch (ArgumentException)
                        {
                            PasswordPrompt passwordDialog = new PasswordPrompt();
                            if (passwordDialog.ShowDialog() == DialogResult.OK)
                            {
                                PdfDocument unlockedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray(), new ParsingOptions { Password = passwordDialog.Password });
                                try { IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen); }
                                finally { unlockedDoc.Dispose(); }
                            }
                            else
                            {
                                throw new InvalidDataException(_cultureHelper.GetLocalizedString("[ReadPdfFileAsync] InvalidDataException #1"));
                            }
                        }
                    }
                }
            }
        }

        private static void IterateInnerPdfFile(ref PdfDocument doc, ref List<string> chunks, int chunkLen)
        {
            var pages = doc.GetPages();
            foreach (var page in pages)
            {
                var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(page.GetWords());
                foreach (var block in blocks)
                    chunks.AddRange(CommonUtils.SplitString(block.Text, chunkLen));
            }
        }

        public string GetRAGContext(string query, int maxTokens)
        {
            if (_fileList.Count == 0) return string.Empty;
            var result = _db.QueryCosineSimilarity(query, _fileList.Count * 10); // 10 results per file
            StringBuilder ragContext = new StringBuilder();
            foreach (var document in result.Documents)
                ragContext.AppendLine(document.DocumentString);
            return CommonUtils.SubstringTokens(ragContext.ToString(), maxTokens);
        }

        // UTILS
        public static AsyncCollectionResult<StreamingChatCompletionUpdate> AskQuestion(SystemChatMessage systemPrompt, IEnumerable<ChatMessage> messages, Word.Range context, float temperature, Word.Document doc = null)
        {
            var chatHistory = ProcessInformation(systemPrompt, messages, context, doc);

            ChatClient client = new ChatClient(ThisAddIn.Model, new ApiKeyCredential(ThisAddIn.ApiKey), ThisAddIn.ClientOptions);

            // https://github.com/ollama/ollama/pull/6504
            return client.CompleteChatStreamingAsync(
                chatHistory,
                new ChatCompletionOptions() { Temperature = temperature * 2 },
                ThisAddIn.CancellationTokenSource.Token
            );
        }
        public static Task<ClientResult<GeneratedImage>> AskQuestionForImage(SystemChatMessage systemPrompt, IEnumerable<ChatMessage> messages, Word.Range context, Word.Document doc = null)
        {
            var chatHistory = ProcessInformation(systemPrompt, messages, context, doc);

            ImageClient client = new ImageClient(ThisAddIn.Model, new ApiKeyCredential(ThisAddIn.ApiKey), ThisAddIn.ClientOptions);

            return client.GenerateImageAsync(
                ModelProperties.ConvertChatHistoryToString(chatHistory),
                new ImageGenerationOptions() { ResponseFormat = GeneratedImageFormat.Bytes },
                ThisAddIn.CancellationTokenSource.Token
            );
        }

        private static List<ChatMessage> ProcessInformation(SystemChatMessage systemPrompt, IEnumerable<ChatMessage> messages, Word.Range context, Word.Document doc = null)
        {
            if (doc == null)
                doc = context.Document;
            string document = context.Text;
            int userPromptLen = GetUserPromptLen(messages);
            ChatMessage lastUserPrompt = messages.Last();

            var constraints = RAGControl.OptimizeConstraint(
                0.8f,
                ThisAddIn.ContextLength,
                CommonUtils.CharToTokenCount(systemPrompt.Content[0].Text.Length + userPromptLen),
                CommonUtils.CharToTokenCount(document.Length)
            );
            if (constraints["document_content_rag"] == 1f)
                document = RAGControl.GetWordDocumentAsRAG(lastUserPrompt.Content[0].Text, context);

            string ragQuery =
                (constraints["rag_context"] == 0f) ? string.Empty : ThisAddIn.AllTaskPanes[doc].Item3.GetRAGContext(lastUserPrompt.Content[0].Text, (int)(ThisAddIn.ContextLength * constraints["rag_context"]));

            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {
                systemPrompt,
                new UserChatMessage($@"{Forge.CultureHelper.GetLocalizedString("(RAGControl.cs) [AskQuestion] chatHistory #1")}\n""{CommonUtils.SubstringTokens(document, (int)(ThisAddIn.ContextLength * constraints["document_content"]))}""")
            };
            if (ragQuery != string.Empty)
                chatHistory.Add(new UserChatMessage($@"{Forge.CultureHelper.GetLocalizedString("(RAGControl.cs) [AskQuestion] chatHistory #2")}\n""{ragQuery}"""));
            chatHistory.AddRange(messages);

            return chatHistory;
        }

        private static int GetUserPromptLen(IEnumerable<ChatMessage> messageList)
        {
            int userPromptLen = 0;
            foreach (var message in messageList)
                userPromptLen += message.Content[0].Text.Length;
            return userPromptLen;
        }

        public static Dictionary<string, float> OptimizeConstraint(float maxPercentage, int contextLength, int promptTokenLen, int documentContentTokenLen)
        {
            Dictionary<string, float> constraintPairs = new();
            if (promptTokenLen >= contextLength * 0.9)
            {
                constraintPairs["rag_context"] = 0f;
                constraintPairs["document_content"] = (float)(maxPercentage * 0.1);
                constraintPairs["document_content_rag"] = (documentContentTokenLen > contextLength * maxPercentage * 0.1) ? 1f : 0f;
            }
            else
            {
                float promptPercentage = (float)promptTokenLen / (float)contextLength;
                constraintPairs["rag_context"] = (float)( (1 - promptPercentage) * maxPercentage * 0.3);
                constraintPairs["document_content"] = (float)((1 - promptPercentage) * maxPercentage * 0.7);
                constraintPairs["document_content_rag"] = (documentContentTokenLen > contextLength * (1 - promptPercentage) * maxPercentage * 0.7) ? 1f : 0f;
            }
            return constraintPairs;
        }

        public static string GetWordDocumentAsRAG(string query, Word.Range context)
        {
            // Get RAG context
            HyperVectorDB.HyperVectorDB DB = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
            var chunks = CommonUtils.SplitString(context.Text, CHUNK_LEN);
            foreach (var chunk in chunks)
                DB.IndexDocument(chunk);

            var result = DB.QueryCosineSimilarity(query, CommonUtils.GetWordPageCount() * 3);

            StringBuilder ragContextBuilder = new StringBuilder(result.Documents.Count);
            foreach (var doc in result.Documents)
                ragContextBuilder.AppendLine(doc.DocumentString);
            return ragContextBuilder.ToString();
        }
    }
}
