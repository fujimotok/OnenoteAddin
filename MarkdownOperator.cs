using ColorCode;
using Markdig;
using Markdig.Renderers;
using Markdig.Renderers.Html;
using Markdig.Renderers.Html.Inlines;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace OnenoteAddin
{
    internal class MarkdownOperator
    {
        public static string ConvertMarkdownToHtml(string markdown)
        {
            using (var writer = new StringWriter())
            {
                var renderer = new HtmlRenderer(writer);

                renderer.ObjectRenderers.RemoveAll(r => r is CodeInlineRenderer);
                renderer.ObjectRenderers.Add(new OneNoteCodeInlineRenderer());

                var originalCodeBlockRenderer = renderer.ObjectRenderers.FindExact<CodeBlockRenderer>();
                renderer.ObjectRenderers.Add(new HighlightedCodeBlockRenderer(originalCodeBlockRenderer));

                if (originalCodeBlockRenderer != null)
                {
                    renderer.ObjectRenderers.Remove(originalCodeBlockRenderer);
                }

                var pipeline = new MarkdownPipelineBuilder()
                    //.UseAdvancedExtensions() // DiaglamがCodeBlockRenderer前提なので各個設定
                    //.UseAlertBlocks()        // ::: alert などで警告・注意・成功などのスタイル付きブロックを作成
                    //.UseAbbreviations()      // *[HTML]: HyperText Markup Language のような略語定義をサポート
                    //.UseAutoIdentifiers()    // 見出しに自動で id を付与（リンク用）
                    //.UseCitations()          // [@doe2020] のような文献引用を可能にする
                    //.UseCustomContainers()   // ::: note や ::: tip のようなカスタムブロックを定義可能
                    //.UseDefinitionLists()    // 定義リスト（用語: 説明）を作成できる
                    .UseEmphasisExtras()       // 打ち消し線（~~text~~）、下付き・上付き、挿入・強調などを追加
                                               //.UseFigures()            // ![alt テキスト](画像URL "キャプションテキスト") 画像にキャプションを付けて <figure> タグで出力
                                               //.UseFooters()            // 文末にフッターを追加できる（[^footer] など）
                                               //.UseFootnotes()          // 脚注を使えるようにする（[^1] など）
                    .UseGridTables()           // 複雑な表（行・列の結合など）を作成できる
                                               //.UseMathematics()        // $x^2$ や $$\int f(x)dx$$ のようなLaTeX数式をサポート（MathJax等と併用）
                                               //.UseMediaLinks()         // YouTubeやVimeoなどのURLを埋め込みメディアに変換（iframeなど）
                    .UsePipeTables()           // `|` を使ったシンプルな表記法の表をサポート
                    .UseListExtras()           // リストにチェックボックスや複雑な構造を追加
                                               //.UseTaskLists()          // チェック付きタスクリスト（- [x] や - [ ]）をサポート
                                               //.UseDiagrams()           // MermaidやGraphvizなどの図表をMarkdown内で記述可能（外部JS連携が必要
                    .UseAutoLinks()            // http://example.com を自動的にリンク化
                                               //.UseGenericAttributes()  // {.class #id} のように任意のHTML属性をMarkdown要素に付与可能
                    .UseSoftlineBreakAsHardlineBreak() // 改行をスペース2個ではなく、改行で行う
                    .Build();


                pipeline.Setup(renderer);

                var doc = Markdown.Parse(markdown, pipeline);

                renderer.Render(doc);
                writer.Flush();

                return writer.ToString();
            }
        }
    }

    internal class OneNoteCodeInlineRenderer : HtmlObjectRenderer<CodeInline>
    {
        protected override void Write(HtmlRenderer renderer, CodeInline obj)
        {
            renderer.Write("<span class='code-block'>");
            renderer.WriteEscape(obj.Content);
            renderer.Write("</span>");
        }
    }

    internal class HighlightedCodeBlockRenderer : HtmlObjectRenderer<CodeBlock>
    {
        private readonly CodeBlockRenderer _underlyingRenderer;

        public HighlightedCodeBlockRenderer(CodeBlockRenderer underlyingRenderer = null)
        {
            _underlyingRenderer = underlyingRenderer ?? new CodeBlockRenderer();
        }

        protected override void Write(HtmlRenderer renderer, CodeBlock obj)
        {
            var fencedCB = obj as FencedCodeBlock;

            if (fencedCB == null)
            {
                // FencedCodeBlock でなければ元のレンダラーに任せる
                _underlyingRenderer.Write(renderer, obj);
                return;
            }

            // cpp:sample.cpp のように言語指定とファイル名がある場合、言語指定のみ抽出
            var langId = "java"; // planetextがないので、デフォルトは java にする
            var info = fencedCB.Info ?? string.Empty;
            var match = Regex.Match(info, @"\b([a-zA-Z0-9_]+)(?=(:|$))");
            if (match.Success)
            {
                langId = match.Groups[1].Value;
            }

            var language = ColorCode.Languages.FindById(langId);

            var lines = fencedCB.Lines.ToString();

            var formatter = new HtmlFormatter();
            var code = formatter.GetHtmlString(lines, language);
            var html = $@"<table><tr class='code-block'><td>```{info}{code}<br>```</td></tr></table>";

            renderer.Write(html);
        }
    }
}
