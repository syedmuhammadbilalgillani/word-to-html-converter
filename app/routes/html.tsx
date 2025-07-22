import { HTMLInlineCSS } from "~/welcome/html-css";
import type { Route } from "./+types/home";

export function meta({}: Route.MetaArgs) {
  return [
    { title: "HTMLInlineCSS" },
    { name: "description", content: "Welcome to React Router!" },
  ];
}

export default function Home() {
  return <HTMLInlineCSS />;
}
