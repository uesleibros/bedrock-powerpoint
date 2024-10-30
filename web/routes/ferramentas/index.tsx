export default function Tools() {
  const tools = [
    {
      id: "discord-rpc",
      name: "Discord RPC",
      description: "Integre o Discord Rich Presence ao VBA, permitindo exibir status personalizados em tempo real.",
      image: "/ferramentas/discord-rpc.png",
    },
    {
      id: "temporizer",
      name: "Temporizer",
      description: "Facilita o controle de tempo e execução de tarefas com temporizadores avançados em VBA.",
      image: "/ferramentas/temporizer.png",
    },
  ];

  return (
    <div class="py-12 px-6">
      <div class="container mx-auto max-w-5xl">
        <div class="flex items-center gap-3 mb-10">
          <a href="/" class="bg-white border rounded-lg shadow-lg p-2 h-[max-content] text-gray-800 font-semibold">
            &larr;
          </a>
          <h2 class="text-4xl mx-auto font-bold">Ferramentas</h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          {tools.map((tool) => (
            <div key={tool.name} class="bg-white border p-6 rounded-lg shadow-lg flex flex-col items-center transition transform hover:-translate-y-1 hover:shadow-2xl">
              <img src={tool.image} alt={`${tool.name} logo`} class="w-[50px] h-[50px] object-contain mb-4" />
              <h3 class="text-xl font-bold text-indigo-700 mb-2">{tool.name}</h3>
              <p class="text-gray-700 text-center">{tool.description}</p>
              <a href={`/ferramentas/documentacao/${tool.id}`} class="bg-white border rounded-lg shadow-lg p-4 text-gray-800 font-semibold mt-5 transition duration-300 transform hover:scale-105">
                Documentação
              </a>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
