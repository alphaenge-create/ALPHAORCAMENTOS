import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // IMPORTANTE: troque "orcacpu" pelo nome exato do seu repositorio no GitHub.
  base: "/orcacpu/",
  build: {
    rollupOptions: {
      output: {
        manualChunks: {
          firebase: ["firebase/app", "firebase/firestore"],
          xlsx: ["xlsx"],
          react: ["react", "react-dom"],
        },
      },
    },
  },
});
