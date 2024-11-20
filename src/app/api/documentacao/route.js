import fs from "fs/promises";
import path from "path";

export async function GET(request) {
  const ferramenta = request.nextUrl.searchParams.get("ferramenta");

  if (!ferramenta)
    return Response.json({ message: "Missing required field (ferramenta).", error: "missing field" }, { status: 401 });

  const filePath = path.resolve(process.cwd(), "src", "docs", `${ferramenta}.md`);

  try {
    await fs.access(filePath, fs.constants.F_OK);

    const content = await fs.readFile(filePath, "utf-8");
    return Response.json({ content }, { status: 200 });
  } catch (err) {
    return Response.json({ message: "File not found.", error: err }, { status: 404 });
  }
}
