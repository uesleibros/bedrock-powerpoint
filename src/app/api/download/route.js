import tools from "@/tools";
import fs from "fs/promises";
import path from "path";

export async function GET(request) {
  const ferramenta = request.nextUrl.searchParams.get("ferramenta");

  if (!ferramenta) {
    return new Response(
      JSON.stringify({ message: "Missing required field (ferramenta).", error: "missing field" }),
      { status: 401 }
    );
  }

  const tool = tools.find((t) => t.id === ferramenta);

  if (!tool) {
    return new Response(
      JSON.stringify({ message: "Ferramenta não encontrada." }),
      { status: 404 }
    );
  }

  const filePath = path.resolve(process.cwd(), "src", "packages", tool.file);

  try {
    await fs.access(filePath, fs.constants.F_OK);

    const fileStats = await fs.stat(filePath);
    const file = await fs.readFile(filePath);

    return new Response(file, {
      headers: {
        "Content-Type": "application/octet-stream",
        "Content-Disposition": `attachment; filename="${tool.file}"`,
        "Content-Length": fileStats.size.toString()
      },
    });
  } catch (error) {
    return new Response(
      JSON.stringify({ message: "Erro ao baixar o pacote.", error: error.message }),
      { status: 500 }
    );
  }
}