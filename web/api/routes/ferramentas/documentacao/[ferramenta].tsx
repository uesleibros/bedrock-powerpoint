import { PageProps } from "$fresh/server";
import { marked } from "https://deno.land/x/marked/mod";
import tools from "../../../tools";

async function loadMarkdown(file: string) {
  const url = new URL(`../../../docs/${file}`, import.meta.url);
  const markdown = await Deno.readTextFile(url);
  return markdown;
}

export const handler = {
  async GET(req: Request, ctx: { params: { ferramenta: string } }) {
    const { ferramenta } = ctx.params;

    if (!ferramenta) {
      return new Response(
        JSON.stringify({ error: "Ferramenta não encontrada." }),
        { status: 404 },
      );
    }

    const tool = tools.find((t) => t.id === ferramenta);

    const markdownFile = `${ferramenta}.md`;
    let content;

    try {
      content = await loadMarkdown(markdownFile);
    } catch (error) {
      content = "# Documentação não encontrada";
    }

    const htmlContent = marked(content);

    return new Response(`
      <!DOCTYPE html>
      <html lang="pt-BR">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${tool.name}</title>
          <link rel="stylesheet" href="/styles.css">
          <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/markdown-css@1.1.0/markdown.css">
          <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.6.0/styles/default.min.css">
          <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.6.0/highlight.min.js"></script>
        </head>
        <body>
          <div class="py-12 px-6">
            <div class="container mx-auto max-w-5xl">
              <div class="flex items-center gap-3 mb-10">
                <a href="/ferramentas" class="bg-white border rounded-lg shadow-lg p-2 h-[max-content] text-gray-800 font-semibold">
                  &larr;
                </a>
                <a href="/api/ferramentas/${ferramenta}/download" class="bg-white border rounded-lg shadow-lg p-2 text-gray-800 font-semibold">
                  Baixar Pacote
                </a>
              </div>
              <div class="pl-[1rem] mx-auto max-w-[800px]">
                <img src="${tool.image}" alt="${tool.name} logo" class="w-[80px] h-[80px] object-contain" />
              </div>
              <div class="prose" style="max-width: 800px;">
                ${htmlContent}
              </div>
            </div>
          </div>
          <script>
            document.addEventListener("DOMContentLoaded", () => {
              hljs.highlightAll();
            });
          </script>
        </body>
      </html>
    `, {
      headers: {
        "Content-Type": "text/html; charset=utf-8",
      },
    });
  },
};
