import { defineConfig } from "vite";
import { resolve } from "path";
import fs from "fs";

export default defineConfig({
  plugins: [
    {
      name: "copy-static-dirs",
      closeBundle() {
        ["js", "css", "assets"].forEach(dir => {
          if (fs.existsSync(dir)) {
            fs.cpSync(dir, resolve("dist", dir), { recursive: true });
          }
        });
        if (fs.existsSync("application.html")) {
          fs.copyFileSync("application.html", resolve("dist", "application.html"));
        }
      }
    }
  ]
});
