import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  base: "/fishrep/",   // <- coloque o nome EXATO do seu repositório GitHub aqui
});
