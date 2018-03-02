using Esprima;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;

namespace CrmIntellisense
{
    public class PositionHelper
    {
        public static bool IsStringLiteral(ICompletionSession session, ITextBuffer mTextBuffer)
        {
            try
            {
                SnapshotPoint caretPosition = session.TextView.Caret.Position.BufferPosition;

                Document document = mTextBuffer.CurrentSnapshot.GetOpenDocumentInCurrentContextWithChanges();

                SyntaxToken token = document.GetSyntaxRootAsync().Result.FindToken(caretPosition);

                return token.IsKind(SyntaxKind.StringLiteralToken);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsLiteralOrTemplate(ICompletionSession session, ITextBuffer mTextBuffer)
        {
            bool isInLiteralOrTemplate = false;

            var bufferPosition = session.TextView.Caret.Position.BufferPosition;
            var snapshot = bufferPosition.Snapshot;
            var position = bufferPosition.Position;
            var line = snapshot.GetLineFromPosition(position);
            int lineNumber = line.LineNumber + 1;
            int column = position - line.Start.Position + 1;

            JToken node = ParseJavaScript(mTextBuffer);
            if (node == null)
                return false;

            WalkNode(node, n =>
            {
                JToken token = n["Type"];
                if (token == null || token.Type != JTokenType.String ||
                    token.Value<string>() != "Literal" && token.Value<string>() != "TemplateLiteral")
                    return;

                var loc = n["Loc"];
                var start = loc?["Start"];
                var startLine = start?["Line"];
                var startColumn = start?["Column"];
                int startLineNumber = startLine.Value<int>();
                int startColumnNumber = startColumn.Value<int>();
                var end = loc?["End"];
                var endLine = end?["Line"];
                var endColumn = end?["Column"];
                int endLineNumber = endLine.Value<int>();
                int endColumnNumber = endColumn.Value<int>();

                // Inside single '' or double "" quotes
                if (token.Value<string>() == "Literal")
                {
                    if (lineNumber == startLineNumber && lineNumber == endLineNumber &&
                        column >= startColumnNumber && column <= endColumnNumber)

                        isInLiteralOrTemplate = true;
                }

                //Inside multi-line text ``
                if (token.Value<string>() == "TemplateLiteral")
                {
                    if (lineNumber < startLineNumber || lineNumber > endLineNumber)
                        return;
                    if (lineNumber == startLineNumber && column < startColumnNumber)
                        return;
                    if (lineNumber == endLineNumber && column > endColumnNumber)
                        return;

                    isInLiteralOrTemplate = true;
                }
            });

            return isInLiteralOrTemplate;
        }

        private static JToken ParseJavaScript(ITextBuffer mTextBuffer)
        {
            try
            {
                var parser = new JavaScriptParser(mTextBuffer.CurrentSnapshot.AsText().ToString(), new ParserOptions
                {
                    Loc = true,
                    Tolerant = true
                });

                var jsAst = parser.ParseProgram();
                var source = JsonConvert.SerializeObject(jsAst);
                JToken node = JToken.Parse(source);

                return node;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void WalkNode(JToken node, Action<JObject> action)
        {
            if (node.Type == JTokenType.Object)
            {
                action((JObject)node);

                foreach (JProperty child in node.Children<JProperty>())
                {
                    WalkNode(child.Value, action);
                }
            }
            else if (node.Type == JTokenType.Array)
            {
                foreach (JToken child in node.Children())
                {
                    WalkNode(child, action);
                }
            }
        }
    }
}