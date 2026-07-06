import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  base: "/ALPHAORCAMENTOS/",
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
