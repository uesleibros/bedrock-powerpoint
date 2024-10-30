import tools from "../../../../tools";

export const handler = {
  async GET(req: Request, ctx: { params: { ferramenta: string } }) {
    const { ferramenta } = ctx.params;

    const tool = tools.find((t) => t.id === ferramenta);

    if (!tool) {
      return new Response(
        JSON.stringify({ error: "Ferramenta não encontrada." }),
        { status: 404 },
      );
    }

    try {
      const filePath = `../../../../packages/${tool.file}`;
      const url = new URL(filePath, import.meta.url);
      const file = await Deno.readFile(url);

      return new Response(file, {
        headers: {
          "Content-Type": "application/octet-stream",
          "Content-Disposition": `attachment; filename="${tool.file}"`,
        },
      });
    } catch (error) {
      console.error("Erro ao ler o arquivo:", error);
      return new Response(
        JSON.stringify({ error: "Erro ao baixar o pacote." }),
        { status: 500 },
      );
    }
  },
};