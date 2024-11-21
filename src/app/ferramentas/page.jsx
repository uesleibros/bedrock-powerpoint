import tools from "@/tools";
import Image from "next/image";
import Link from "next/link";

export default async function Ferramentas() {
	return (
    <div className="w-full h-full min-h-screen h-full antialiased w-full bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:30px_30px]">
  		<div className="py-12 px-6">
        <div className="container mx-auto max-w-5xl">
          <div className="flex items-center gap-3 mb-10">
            <Link href="/" className="bg-white border rounded-lg shadow-lg p-2 h-[max-content] text-gray-800 font-semibold">
              &larr;
            </Link>
            <h2 className="text-4xl mx-auto font-bold">Ferramentas</h2>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
            {tools.map((tool) => (
              <div key={tool.name} className="relative bg-white border-b p-6 rounded-lg shadow-lg flex flex-col items-start transition transform hover:-translate-y-1 hover:shadow-2xl">
                <div className="absolute top-0 left-0 w-full h-[100px] bg-blue-600 z-0 rounded-tl-lg rounded-tr-lg border-t border-l border-r"></div>
                <div className="mt-[40px] z-10 relative">
                  <Image src={tool.image} alt={`${tool.name} logo`} width={1000} height={1000} quality={100} className="w-[60px] h-[60px] pointer-events-none select-none object-contain mb-4 rounded-lg" />
                  <div className="my-1 flex flex-wrap items-center gap-2">
                    {tool.tags.map((tag, index) => (
                      <p key={index} className="bg-blue-600 select-none text-xs text-white font-bold rounded-full shadow-sm py-1 px-4">{ tag }</p>
                    ))}
                  </div>
                  <h3 className="text-xl font-bold bg-gradient-to-r from-black to-gray-900 bg-clip-text text-transparent mb-2">{tool.name}</h3>
                  <p className="text-gray-700 mb-5">{tool.description}</p>
                  <Link href={`/ferramentas/documentacao/${tool.id}`} className="inline-block bg-white border rounded-lg shadow-lg p-4 w-full text-center text-gray-800 font-semibold mt-auto transition duration-300 transform hover:scale-105">
                    Documentação
                  </Link>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
	);
}