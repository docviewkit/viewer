import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { createCommercialApp } from "./app.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
const publicDir = join(root, "public");
const host = process.env.HOST || "127.0.0.1";
const port = Number(process.env.PORT || 4310);

const app = createCommercialApp({
  publicDir,
  immutableSdkAssets: process.env.NODE_ENV === "production",
  indexNowKey: process.env.DOCVIEWKIT_INDEXNOW_KEY || "",
  siteVerification: {
    google: process.env.DOCVIEWKIT_GOOGLE_SITE_VERIFICATION || "",
    bing: process.env.DOCVIEWKIT_BING_SITE_VERIFICATION || "",
  },
});
app.server.listen(port, host, () => {
  console.log(`DocViewKit website: http://${host}:${port}`);
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, async () => {
    await app.close();
    process.exit(0);
  });
}
