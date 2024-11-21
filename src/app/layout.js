import { Noto_Sans } from "next/font/google";
import "./globals.css";

export const metadata = {
  title: "Bedrock PowerPoint",
  description: "Uma equipe de jogos em PowerPoint.",
};

const font = Noto_Sans({ subsets: ["latin"], display: "swap" })

export default function RootLayout({ children }) {
  return (
    <html lang="pt-BR">
      <body
        className={`${font.className} antialiased`}
      >
        {children}
      </body>
    </html>
  );
}
