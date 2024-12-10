using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using OpenAI.Chat;

namespace TextForge
{
    public partial class GenerateUserControl : UserControl
    {
        public static readonly CultureLocalizationHelper CultureHelper = new CultureLocalizationHelper("TextForge.GenerateUserControl", typeof(GenerateUserControl).Assembly);

        public GenerateUserControl()
        {
            try
            {
                InitializeComponent();
                MatchScrollBarTemperature(); // for floating-point localization
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void GenerateButton_Click(object sender, EventArgs e)
        {
            try
            {
                string textBoxContent = this.PromptTextBox.Text;
                if (textBoxContent.Length == 0)
                    throw new EmptyTextBoxException(CultureHelper.GetLocalizedString("[GenerateButton_Click] TextBoxEmptyException #1"));

                /*
                 * So, If the user changes the selection carot in Word after clicking "generate" (bc it takes so long to generate text).
                 * Then, it won't affect where the text is placed.
                 */
                var rangeBeforeChat = Globals.ThisAddIn.Application.Selection.Range;
                var docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();

                // Clear any selected text by the user
                if (rangeBeforeChat.End - rangeBeforeChat.Start > 0)
                    rangeBeforeChat.Delete();

                if (ModelProperties.IsImageModel(ThisAddIn.Model))
                {
                    var streamingAnswer = RAGControl.AskQuestionForImage(
                        new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]),
                        new List<UserChatMessage> { new UserChatMessage(textBoxContent) },
                        docRange
                    );
                    await Forge.AddStreamingImageContentToRange(streamingAnswer, rangeBeforeChat);
                }
                else
                {
                    var streamingAnswer = RAGControl.AskQuestion(
                        new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(GenerateUserControl.cs) _systemPrompt"]),
                        new List<UserChatMessage> { new UserChatMessage(textBoxContent) },
                        docRange,
                        GetTemperature()
                    );
                    await Forge.AddStreamingChatContentToRange(streamingAnswer, rangeBeforeChat);
                }

                Globals.ThisAddIn.Application.Selection.SetRange(rangeBeforeChat.Start, rangeBeforeChat.End);
            }
            catch (EmptyTextBoxException ex)
            {
                CommonUtils.DisplayInformation(ex);
            }
            catch (OperationCanceledException ex)
            {
                CommonUtils.DisplayWarning(ex);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void PromptTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control && e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    this.GenerateButton.PerformClick();
                }
                else if (e.Control && e.KeyCode == Keys.Back)
                {
                    e.SuppressKeyPress = true;
                    int cursorPosition = this.PromptTextBox.SelectionStart;
                    string text = this.PromptTextBox.Text;

                    // Handle multiple trailing spaces
                    while (cursorPosition > 0 && text[cursorPosition - 1] == ' ')
                    {
                        cursorPosition--;
                    }

                    text = text.TrimEnd();

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        this.PromptTextBox.Clear();
                        this.PromptTextBox.SelectionStart = 0;
                    }
                    else
                    {
                        int lastSpaceIndex = text.LastIndexOf(' ', cursorPosition - 1);
                        if (lastSpaceIndex != -1)
                        {
                            // Retain a space after deletion
                            this.PromptTextBox.Text = text.Remove(lastSpaceIndex + 1, cursorPosition - lastSpaceIndex - 1);
                            this.PromptTextBox.SelectionStart = lastSpaceIndex + 1;
                        }
                        else
                        {
                            this.PromptTextBox.Text = text.Remove(0, cursorPosition);
                            this.PromptTextBox.SelectionStart = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void TemperatureTrackBar_Scroll(object sender, EventArgs e)
        {
            try
            {
                MatchScrollBarTemperature();
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void MatchScrollBarTemperature()
        {
            this.TemperatureValueLabel.Text = GetTemperature().ToString("0.0", Thread.CurrentThread.CurrentUICulture);
        }

        private float GetTemperature()
        {
            return this.TemperatureTrackBar.Value / 10f;
        }
    }

    public class EmptyTextBoxException : ArgumentException
    {
        public EmptyTextBoxException(string message) : base(message) { }
    }
}
