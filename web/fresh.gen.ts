// DO NOT EDIT. This file is generated by Fresh.
// This file SHOULD be checked into source version control.
// This file is automatically updated during development when running `dev.ts`.

import * as $_404 from "./routes/_404.tsx";
import * as $_app from "./routes/_app.tsx";
import * as $api_ferramentas_ferramenta_download from "./routes/api/ferramentas/[ferramenta]/download.ts";
import * as $ferramentas_documentacao_ferramenta_ from "./routes/ferramentas/documentacao/[ferramenta].tsx";
import * as $ferramentas_index from "./routes/ferramentas/index.tsx";
import * as $index from "./routes/index.tsx";

import type { Manifest } from "$fresh/server.ts";

const manifest = {
  routes: {
    "./routes/_404.tsx": $_404,
    "./routes/_app.tsx": $_app,
    "./routes/api/ferramentas/[ferramenta]/download.ts":
      $api_ferramentas_ferramenta_download,
    "./routes/ferramentas/documentacao/[ferramenta].tsx":
      $ferramentas_documentacao_ferramenta_,
    "./routes/ferramentas/index.tsx": $ferramentas_index,
    "./routes/index.tsx": $index,
  },
  islands: {},
  baseUrl: import.meta.url,
} satisfies Manifest;

export default manifest;
