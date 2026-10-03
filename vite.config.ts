import { defineConfig } from "vitest/config";

export default defineConfig({
  // Relative asset paths so GitHub Pages works from a repository subpath.
  base: "./",
  test: {
    environment: "node",
    include: ["tests/**/*.test.ts"],
    setupFiles: ["tests/setup.ts"],
  },
});
