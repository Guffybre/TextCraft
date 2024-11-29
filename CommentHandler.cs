using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using OpenAI.Chat;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    internal class CommentHandler
    {
        private static int _prevNumComments = 0;
        private static bool _isDraftingComment = false;

        public static async void Document_CommentsEventHandler(Word.Selection selection)
        {
            try
            {
                // For preventing unnecessary iteration of this function every time something changes in Word.
                int numComments = Globals.ThisAddIn.Application.ActiveDocument.Comments.Count;
                if (numComments == _prevNumComments) return;

                if (await AICommentReplyTask())
                    numComments++;

                if (await UserMentionTask())
                    numComments++;

                _prevNumComments = numComments;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static async Task<bool> AICommentReplyTask()
        {
            var comments = GetUnansweredAIComments(Globals.ThisAddIn.Application.ActiveDocument.Comments);
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (var comment in comments)
            {
                List<ChatMessage> chatHistory = new List<ChatMessage>() {
                    new UserChatMessage($@"{Forge.CultureHelper.GetLocalizedString("[Review] chatHistory #1")}\n""{CommonUtils.SubstringTokens(comment.Range.Text, (int)(ThisAddIn.ContextLength * 0.2))}"""),
                    new UserChatMessage(Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #2"))
                };
                chatHistory.AddRange(GetCommentMessages(comment));
                chatHistory.Add(new UserChatMessage(@$"Text in Focus:\n""{comment.Scope.Text}"""));

                try
                {
                    if (_isDraftingComment) return false; // TODO: is this really necessary?
                    _isDraftingComment = true;
                    
                    await AddComment(
                        comment.Replies,
                        comment.Range,
                        RAGControl.AskQuestion(Forge.CommentSystemPrompt, chatHistory, Globals.ThisAddIn.Application.ActiveDocument.Range(), doc)
                    );

                    _isDraftingComment = false;
                    return true;
                }
                catch (OperationCanceledException ex)
                {
                    CommonUtils.DisplayWarning(ex);
                }
            }
            return false;
        }

        private static async Task<bool> UserMentionTask()
        {
            var comments = GetUnansweredMentionedComments(Globals.ThisAddIn.Application.ActiveDocument.Comments);
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (var comment in comments)
            {
                List<ChatMessage> chatHistory = new List<ChatMessage>();
                chatHistory.AddRange(GetCommentMessagesWithoutMention(comment));
                chatHistory.Add(new UserChatMessage(@$"Text in Focus:\n""{comment.Scope.Text}"""));
                try
                {
                    if (_isDraftingComment) return false; // TODO: is this really necessary?
                    _isDraftingComment = true;
                    
                    await AddComment(
                        comment.Replies,
                        comment.Range,
                        RAGControl.AskQuestion(
                            new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(CommentHandler.cs) [AIUserMentionTask] UserMentionSystemPrompt"]),
                            chatHistory,
                            Globals.ThisAddIn.Application.ActiveDocument.Range(),
                            doc
                        )
                    );
                    
                    _isDraftingComment = false;
                    return true;
                }
                catch (OperationCanceledException ex)
                {
                    CommonUtils.DisplayWarning(ex);
                }
            }
            return false;
        }

        private static IEnumerable<ChatMessage> GetCommentMessagesWithoutMention(Comment parentComment)
        {
            string modelName = $"@{ThisAddIn.Model}";

            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {
                new UserChatMessage(GetCleanedCommentText(parentComment, modelName))
            };

            Comments childrenComments = parentComment.Replies; // Includes parent comment
            for (int i = 1; i <= childrenComments.Count; i++)
            {
                var comment = childrenComments[i];
                string cleanText = GetCleanedCommentText(parentComment, modelName);
                chatHistory.Add(
                    (i % 2 == 1) ? new AssistantChatMessage(cleanText) : new UserChatMessage(cleanText)
                );
            }

            return chatHistory;
        }

        private static string GetCleanedCommentText(Comment c, string modelName)
        {
            string commentText = c.Range.Text;
            return commentText.Contains(modelName) ? commentText.Remove(commentText.IndexOf(modelName), modelName.Length).TrimStart() : commentText;
        }

        // Converts Word Comment object into a list of ChatMessage that can be fed into the OpenAI API
        private static IEnumerable<ChatMessage> GetCommentMessages(Comment parentComment)
        {
            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {
                new UserChatMessage(parentComment.Range.Text)
            };
            
            Comments childrenComments = parentComment.Replies;
            for (int i = 1; i <= childrenComments.Count; i++)
            {
                var comment = childrenComments[i];
                chatHistory.Add(
                    (i % 2 == 1) ? new AssistantChatMessage(comment.Range.Text) : new UserChatMessage(comment.Range.Text)
                );
            }

            return chatHistory;
        }

        // Checks if the user mentions the AI with '@' character. Example: "@qwen2.5:1.5b"
        private static IEnumerable<Comment> GetUnansweredMentionedComments(Comments allComments)
        {
            List<Comment> comments = new List<Comment>();
            foreach (Comment c in allComments)
                if (
                    c.Ancestor == null &&
                    ( c.Range.Text.Contains($"@{ThisAddIn.Model}") ? ( (c.Replies.Count == 0) || (c.Replies.Count > 0 && c.Replies[c.Replies.Count].Author != ThisAddIn.Model) ) : AreRepliesUnbalanced(c.Replies) )
                )
                    comments.Add(c);

            return comments;
        }

        private static bool AreRepliesUnbalanced(Comments replies)
        {
            int userMentionCount = GetCommentMentionCount($"@{ThisAddIn.Model}", replies);
            int aiAnswerCount = GetCommentAuthorCount(ThisAddIn.Model, replies);
            return (userMentionCount > aiAnswerCount);
        }

        private static int GetCommentMentionCount(string mention, Comments comments)
        {
            int count = 0;
            for (int i = 1; i <= comments.Count; i++)
                if (comments[i].Range.Text != null && comments[i].Range.Text.Contains(mention)) count++;
            return count;
        }

        private static int GetCommentAuthorCount(string author, Comments comments)
        {
            int count = 0;
            for (int i = 1; i <= comments.Count; i++)
                if (comments[i].Author == author) count++;
            return count;
        }

        // Checks replies to comments generated by "Writing Tools->Review" action.
        private static IEnumerable<Comment> GetUnansweredAIComments(Comments allComments)
        {
            List<Comment> comments = new List<Comment>();
            foreach (Comment c in allComments)
                if (c.Ancestor == null &&
                    c.Author == ThisAddIn.Model &&
                    ( c.Replies.Count > 0 && c.Replies[c.Replies.Count].Author != ThisAddIn.Model )
                    )
                    comments.Add(c);

            return comments;
        }

        public static async Task AddComment(Comments comments, Range range, AsyncCollectionResult<StreamingChatCompletionUpdate> streamingContent)
        {
            Word.Comment c = comments.Add(range, string.Empty);
            c.Author = ThisAddIn.Model;
            Word.Range commentRange = c.Range.Duplicate; // Duplicate the range to work with

            StringBuilder comment = new StringBuilder();
            // Run the comment generation in a background thread

            await Task.Run(async () =>
            {
                Forge.CancelButtonVisibility(true);
                try
                {
                    await foreach (var update in streamingContent.WithCancellation(ThisAddIn.CancellationTokenSource.Token))
                    {
                        if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                            break;
                        foreach (var content in update.ContentUpdate)
                        {
                            commentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd); // Move to the end of the range
                            commentRange.Text = content.Text; // Append new text
                            commentRange = c.Range.Duplicate; // Update the range to include the new text
                            comment.Append(content.Text);
                        }
                    }
                } 
                finally
                {
                    Forge.CancelButtonVisibility(false);
                }
                c.Range.Text = WordMarkdown.RemoveMarkdownSyntax(comment.ToString());
            });
        }
    }
}
