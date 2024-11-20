import Link from "next/link";

export default function Home() {
  return (
    <div className="min-h-screen bg-gradient-to-b text-gray-800 py-12 px-6">
      <div className="container mx-auto max-w-3xl text-center">
        <h1 className="font-extrabold text-6xl mb-6">
          Bedrock
        </h1>
        <p className="text-2xl font-light mb-8">
          Ferramentas inovadoras para desenvolvedores VBA em qualquer área.
        </p>
        <p className="text-lg mb-8 leading-relaxed">
          Nossa missão é transformar o VBA em uma plataforma versátil e poderosa. Seja desenvolvendo jogos, sistemas operacionais, ou qualquer outra aplicação, temos as ferramentas e a expertise para facilitar seu trabalho e expandir o potencial do VBA.
        </p>
        <div className="transition w-[max-content] mx-auto duration-300 transform hover:-translate-y-1">
          <Link href="/ferramentas" className="bg-white border rounded-lg shadow-lg p-4 text-gray-800 font-semibold">
            Ver Ferramentas
          </Link>
        </div>
      </div>

      <div id="features" className="container mx-auto max-w-4xl mt-20 px-6">
        <h2 className="text-center text-3xl font-bold mb-8">O que fazemos</h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          <div className="bg-white border text-gray-800 p-6 rounded-lg shadow-lg transform transition-transform hover:scale-105">
            <h3 className="font-bold text-xl flex items-center gap-2">
              🛠️ Ferramentas
            </h3>
            <p className="mt-2">Diversas ferramentas úteis em VBA, desenvolvidas para agilizar o trabalho dos desenvolvedores.</p>
          </div>
          <div className="bg-white border text-gray-800 p-6 rounded-lg shadow-lg transform transition-transform hover:scale-105">
            <h3 className="font-bold text-xl flex items-center gap-2">
              💻 Linguagens
            </h3>
            <p className="mt-2">Desenvolvimento de linguagens baseadas em VBA, criando novas possibilidades para programação.</p>
          </div>
          <div className="bg-white border text-gray-800 p-6 rounded-lg shadow-lg transform transition-transform hover:scale-105">
            <h3 className="font-bold text-xl flex items-center gap-2">
              🧰 Recursos
            </h3>
            <p className="mt-2">Desenvolvemos soluções específicas para membros da comunidade com códigos simplificados e explicados.</p>
          </div>
          <div className="bg-white border text-gray-800 p-6 rounded-lg shadow-lg transform transition-transform hover:scale-105">
            <h3 className="font-bold text-xl flex items-center gap-2">
              🙋‍ Suporte
            </h3>
            <p className="mt-2">Ajudamos a tirar dúvidas sobre problemas com o VBA.</p>
          </div>
        </div>
      </div>

      <div className="container mx-auto max-w-4xl mt-20 px-6">
        <h2 className="text-center text-3xl font-bold mb-8">O que nossos clientes dizem</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          {[
            { text: "As ferramentas da Bedrock simplificaram o desenvolvimento VBA e elevaram a qualidade dos nossos projetos!", author: "Erickssen, CEO da Erilab" },
            { text: "A Bedrock trouxe soluções VBA que realmente melhoraram nosso fluxo de trabalho e nos fizeram economizar muito tempo.", author: "Arfur, Desenvolvedor de jogos" },
            { text: "Eu amei o fato que a Bedrock consegue transformar qualquer tarefa complexa em algo bobo de simples.", author: "Figames, Criador de Snowland" },
            { text: "Cara, eu precisava muito de uma função de Wait que permitisse trabalhar com várias ao mesmo tempo, Bedrock proporcionou a melhor solução de todas, rs.", author: "Primagi, Criador de Wendel" },
            { text: "Oi.", author: "Fabinho, Internauta" }
          ].map((testimonial, index) => (
            <div key={index} className="bg-white border text-gray-800 h-auto p-6 rounded-lg shadow-lg transform transition-transform hover:scale-105">
              <p className="italic">&quot;{testimonial.text}&quot;</p>
              <p className="text-right mt-4 font-semibold">- {testimonial.author}</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
