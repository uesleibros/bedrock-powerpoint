import { marked } from "marked";
import { headers } from "next/headers";
import Image from "next/image";
import Link from "next/link";
import ContentPostCode from "@/components/ContentPostCode";
import tools from "@/tools";

export async function generateMetadata({ params }) {
  const { ferramenta } = await params;

  const headersList = await headers();
  const host = headersList.get("host");
  const forwardedProto = headersList.get("x-forwarded-proto");

  const req = await fetch(`${forwardedProto}://${host}/api/documentacao?ferramenta=${encodeURIComponent(ferramenta)}`);
  const body = await req.json();

  const ferramentaObjeto = tools.find(tool => tool.id === ferramenta);

  if (!ferramentaObjeto) {
    return {
      title: "Ferramenta não encontrada",
      description: "A ferramenta solicitada não foi encontrada.",
      openGraph: {
        title: "Ferramenta não encontrada",
        description: "Não conseguimos encontrar a ferramenta que você procura.",
        images: [],
        type: "website",
      },
      twitter: {
        card: "summary_large_image",
        title: "Ferramenta não encontrada",
        description: "Não conseguimos encontrar a ferramenta que você procura.",
        images: [],
      },
    };
  }

  return {
    title: `${ferramentaObjeto.name} - Documentação`,
    description: `A documentação oficial da ferramenta ${ferramentaObjeto.name}.`,
    openGraph: {
      title: ferramentaObjeto.name,
      description: `Descubra como utilizar a ferramenta ${ferramentaObjeto.name} com a nossa documentação oficial.`,
      images: [ferramentaObjeto.image],
      type: "website",
    },
    twitter: {
      card: "summary_large_image",
      title: ferramentaObjeto.name,
      description: `A documentação oficial da ferramenta ${ferramentaObjeto.name}.`,
      images: [ferramentaObjeto.image],
    },
  };
}

export default async function DocumentacaoFerramenta({ params }) {
  const { ferramenta } = await params;

  const headersList = await headers();
  const host = headersList.get("host");
  const forwardedProto = headersList.get("x-forwarded-proto");

  const req = await fetch(`${forwardedProto}://${host}/api/documentacao?ferramenta=${encodeURIComponent(ferramenta)}`);
  const body = await req.json();

  const ferramentaObjeto = tools.find(tool => tool.id === ferramenta);

  if (!ferramentaObjeto) {
    return (
      <div className="py-12 px-6">
        <div className="container mx-auto max-w-5xl">
          <div className="flex items-center gap-3 mb-10">
            <Link href="/ferramentas" className="bg-white border rounded-lg shadow-lg p-2 h-[max-content] text-gray-800 font-semibold">
              &larr;
            </Link>
          </div>
          <h1 className="text-xl font-semibold text-gray-800">Ferramenta não encontrada</h1>
          <p className="text-gray-600">A ferramenta solicitada não existe ou foi removida.</p>
        </div>
      </div>
    );
  }

  const markdown = body.content ? marked(body.content) : null;

  return (
    <div className="py-12 px-6">
      <div className="container mx-auto max-w-5xl">
        <div className="flex items-center gap-3 mb-10">
          <Link href="/ferramentas" className="bg-white border rounded-lg shadow-lg p-2 h-[max-content] text-gray-800 font-semibold">
            &larr;
          </Link>
          <a href={`/api/download?ferramenta=${ferramenta}`} className="bg-white border rounded-lg shadow-lg p-2 text-gray-800 font-semibold">
            Baixar Pacote
          </a>
        </div>
        <div className="pl-[1rem] mx-auto max-w-[800px]">
          <Image src={ferramentaObjeto.image} alt={`${ferramentaObjeto.name} logo`} width={1000} height={1000} quality={100} className="w-[80px] h-[80px] pointer-events-none select-none object-contain rounded-lg" />
        </div>
        <div className="pl-[1rem] mx-auto max-w-[800px] mt-5 flex flex-wrap items-center gap-2">
        	{ferramentaObjeto.tags.map((tag, index) => (
        		<p key={index} className="bg-blue-600 select-none text-xs text-white font-bold rounded-full shadow-sm py-1 px-4">{ tag }</p>
        	))}
        </div>
        <ContentPostCode htmlContent={markdown} />
      </div>
    </div>
  );
}