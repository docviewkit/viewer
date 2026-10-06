import { getIndexablePaths } from "../src/docs.mjs";

const key = process.env.DOCVIEWKIT_INDEXNOW_KEY || "";
if (!/^[A-Za-z0-9_-]{8,128}$/u.test(key)) {
  throw new Error("Set DOCVIEWKIT_INDEXNOW_KEY to the key served by the DocViewKit deployment.");
}

const host = "docviewkit.com";
const response = await fetch("https://api.indexnow.org/indexnow", {
  method: "POST",
  headers: { "content-type": "application/json; charset=utf-8" },
  body: JSON.stringify({
    host,
    key,
    keyLocation: `https://${host}/${key}.txt`,
    urlList: getIndexablePaths().map((path) => `https://${host}${path}`),
  }),
});

if (!response.ok) {
  throw new Error(`IndexNow submission failed: ${response.status} ${await response.text()}`);
}

console.log(`Submitted ${getIndexablePaths().length} canonical URLs to IndexNow.`);
