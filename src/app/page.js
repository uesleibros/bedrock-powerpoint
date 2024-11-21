import Link from "next/link";

export default function Home() {
  return (
    <div className="min-h-screen bg-white text-gray-900 antialiased h-full w-full bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:30px_30px]">
      {/* Header Section */}
      <div className="container mx-auto max-w-6xl py-20 px-6 text-center">
        <h1 className="text-7xl font-extrabold tracking-tight bg-gradient-to-r from-black to-gray-900 bg-clip-text text-transparent">
          Bedrock
        </h1>
        <p className="text-xl mt-6 text-gray-600">
          Ferramentas inovadoras para desenvolvedores VBA em qualquer área.
        </p>
        <p className="mt-4 text-lg text-gray-500 leading-relaxed max-w-3xl mx-auto">
          Nossa missão é transformar o VBA em uma plataforma versátil e poderosa. Seja desenvolvendo jogos, sistemas operacionais, ou qualquer outra aplicação, temos as ferramentas e a expertise para facilitar seu trabalho e expandir o potencial do VBA.
        </p>
        <div className="mt-8 animate-slideUp">
          <Link
            href="/ferramentas"
            className="inline-block bg-black text-white px-6 py-3 rounded-md text-lg font-medium hover:bg-gray-900 transition shadow-lg hover:shadow-2xl"
          >
            Ver Ferramentas
          </Link>
        </div>
      </div>

      {/* Features Section */}
      <div className="container mx-auto max-w-6xl py-20 px-6">
        <h2 className="text-3xl font-bold text-center mb-12 text-gray-800">
          O que fazemos
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          {[
            {
              title: "🛠️ Ferramentas",
              text: "Diversas ferramentas úteis em VBA, desenvolvidas para agilizar o trabalho dos desenvolvedores.",
            },
            {
              title: "💻 Linguagens",
              text: "Desenvolvimento de linguagens baseadas em VBA, criando novas possibilidades para programação.",
            },
            {
              title: "🧰 Recursos",
              text: "Desenvolvemos soluções específicas para membros da comunidade com códigos simplificados e explicados.",
            },
            {
              title: "🙋‍ Suporte",
              text: "Ajudamos a tirar dúvidas sobre problemas com o VBA.",
            },
          ].map((feature, index) => (
            <div
              key={index}
              className="bg-gray-50 border border-gray-200 p-8 rounded-xl shadow-lg hover:shadow-xl transition-transform hover:-translate-y-1"
            >
              <h3 className="font-semibold text-lg mb-3 text-gray-800">
                {feature.title}
              </h3>
              <p className="text-gray-600">{feature.text}</p>
            </div>
          ))}
        </div>
      </div>

      {/* Statistics Section */}
      <div className="container mx-auto max-w-6xl py-20 px-6 bg-gray-50 rounded-xl shadow-lg">
        <h2 className="text-3xl font-bold text-center mb-12 text-gray-800">
          Estatísticas que impressionam
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8 text-center">
          {[
            { value: "1+", label: "Linhas de código otimizadas" },
            { value: "1+", label: "Projetos entregues" },
            { value: "1.2%", label: "Taxa de satisfação" },
            { value: "1%", label: "Compromisso com a inovação" },
          ].map((stat, index) => (
            <div key={index} className="flex flex-col items-center">
              <div className="text-6xl font-extrabold text-black">
                {stat.value}
              </div>
              <p className="text-gray-600 mt-2">{stat.label}</p>
            </div>
          ))}
        </div>
      </div>

      {/* Testimonials Section */}
      <div className="container mx-auto max-w-6xl py-20 px-6">
        <h2 className="text-3xl font-bold text-center mb-12 text-gray-800">
          O que nossos clientes dizem
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          {[
            {
              text: "As ferramentas da Bedrock simplificaram o desenvolvimento VBA e elevaram a qualidade dos nossos projetos!",
              author: "Erickssen, CEO da Erilab",
            },
            {
              text: "A Bedrock trouxe soluções VBA que realmente melhoraram nosso fluxo de trabalho e nos fizeram economizar muito tempo.",
              author: "Arfur, Desenvolvedor de jogos",
            },
            {
              text: "Eu amei o fato que a Bedrock consegue transformar qualquer tarefa complexa em algo bobo de simples.",
              author: "Figames, Criador de Snowland",
            },
            {
              text: "Cara, eu precisava muito de uma função de Wait que permitisse trabalhar com várias ao mesmo tempo, Bedrock proporcionou a melhor solução de todas, rs.",
              author: "Primagi, Criador de Wendel",
            },
            { text: "Oi.", author: "Fabinho, Internauta" },
            { text: "", author: "Gabb, CEO da Basement 5.0" }
          ].map((testimonial, index) => (
            <div
              key={index}
              className="bg-gray-50 border border-gray-200 p-6 rounded-xl shadow-lg hover:shadow-xl transition-transform hover:-translate-y-1"
            >
              <p className="italic text-gray-700">&quot;{testimonial.text}&quot;</p>
              <p className="text-right mt-4 font-medium text-gray-900">
                - {testimonial.author}
              </p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
