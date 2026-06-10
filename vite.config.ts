import { existsSync, readFileSync } from "node:fs";
import type { ServerOptions } from "node:https";
import { homedir } from "node:os";
import { join } from "node:path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const getExistingOfficeDevCertOptions = (): ServerOptions | undefined => {
  const certDirectory = join(homedir(), ".office-addin-dev-certs");
  const caPath = join(certDirectory, "ca.crt");
  const certPath = join(certDirectory, "localhost.crt");
  const keyPath = join(certDirectory, "localhost.key");

  if (![caPath, certPath, keyPath].every(existsSync)) {
    return undefined;
  }

  return {
    ca: readFileSync(caPath),
    cert: readFileSync(certPath),
    key: readFileSync(keyPath)
  };
};

export default defineConfig(async ({ command, mode }) => {
  const config = {
    // Required for SharePoint/static-folder hosting: built JS/CSS links must be
    // relative to taskpane.html, not absolute from the SharePoint domain root.
    base: "./",
    plugins: [react()],
    publicDir: "public",
    server: {
      host: "0.0.0.0",
      port: 3000,
      strictPort: false
    },
    build: {
      outDir: "dist",
      rollupOptions: {
        input: "src/taskpane/taskpane.html"
      }
    }
  };

  if (command === "serve" && mode !== "test") {
    // Reuse the already trusted Office dev cert to avoid rerunning install.ps1 on
    // every dev-server start, which can be blocked by managed Windows policies.
    const existingHttpsOptions = getExistingOfficeDevCertOptions();

    if (existingHttpsOptions) {
      return {
        ...config,
        server: {
          ...config.server,
          https: existingHttpsOptions
        }
      };
    }

    const devCerts = await import("office-addin-dev-certs");

    return {
      ...config,
      server: {
        ...config.server,
        https: await devCerts.getHttpsServerOptions()
      }
    };
  }

  return config;
});
