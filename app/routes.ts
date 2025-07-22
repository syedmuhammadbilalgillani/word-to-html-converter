import { type RouteConfig, index, route } from "@react-router/dev/routes";

export default [
  index("routes/home.tsx"),
  route("/html", "routes/html.tsx"),
  route("/only-html", "routes/only-html.tsx"),
] satisfies RouteConfig;
