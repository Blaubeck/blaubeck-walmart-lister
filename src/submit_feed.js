// Submit the MP_ITEM feed to Walmart. Run AFTER updating data/upcs.json with Monday UPCs.
//
// Usage:
//   1. Fill data/upcs.json with the 10 new UPCs from GS1 UK
//   2. node src/build_feed.py (regenerates feed with real UPCs)
//   3. WALMART_CLIENT_ID=... WALMART_CLIENT_SECRET=... node src/submit_feed.js
//
// The feed is sent as multipart/form-data per Walmart spec.

const https = require("https");
const fs = require("fs");
const crypto = require("crypto");
const { URLSearchParams } = require("url");

function uuid() {
  return crypto.randomUUID();
}

function fetchRaw(url, options = {}) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, options, (res) => {
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => resolve({ status: res.statusCode, body: Buffer.concat(chunks) }));
    });
    req.on("error", reject);
    if (options.body) req.write(options.body);
    req.end();
  });
}

async function getToken() {
  const params = new URLSearchParams();
  params.append("grant_type", "client_credentials");
  const r = await fetchRaw("https://marketplace.walmartapis.com/v3/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Accept: "application/json",
      "WM_SVC.NAME": "Walmart Marketplace",
      "WM_QOS.CORRELATION_ID": uuid(),
      Authorization:
        "Basic " +
        Buffer.from(
          process.env.WALMART_CLIENT_ID + ":" + process.env.WALMART_CLIENT_SECRET
        ).toString("base64"),
    },
    body: params.toString(),
  });
  return JSON.parse(r.body.toString()).access_token;
}

function buildMultipart(feedJson) {
  const boundary = "----walmart" + Math.random().toString(36).slice(2);
  const lines = [];
  lines.push(`--${boundary}`);
  lines.push('Content-Disposition: form-data; name="file"; filename="feed.json"');
  lines.push("Content-Type: application/json");
  lines.push("");
  lines.push(feedJson);
  lines.push(`--${boundary}--`);
  lines.push("");
  return { body: lines.join("\r\n"), boundary };
}

async function submit() {
  const token = await getToken();
  const feedJson = fs.readFileSync("data/walmart_feed.json", "utf-8");

  // Sanity check: any placeholders left?
  if (feedJson.indexOf("PLACEHOLDER") !== -1) {
    console.error("ABORT: feed contains UPC placeholders. Fill data/upcs.json and rerun build_feed.py first.");
    process.exit(1);
  }

  const { body, boundary } = buildMultipart(feedJson);

  const res = await fetchRaw(
    "https://marketplace.walmartapis.com/v3/feeds?feedType=MP_ITEM",
    {
      method: "POST",
      headers: {
        Accept: "application/json",
        "Content-Type": "multipart/form-data; boundary=" + boundary,
        "Content-Length": Buffer.byteLength(body),
        "WM_SEC.ACCESS_TOKEN": token,
        "WM_SVC.NAME": "Walmart Marketplace",
        "WM_QOS.CORRELATION_ID": uuid(),
      },
      body,
    }
  );

  console.log("Submit status:", res.status);
  console.log("Response:", res.body.toString());
  if (res.status >= 400) process.exit(1);

  const resp = JSON.parse(res.body.toString());
  const feedId = resp.feedId;
  console.log("FeedId:", feedId);

  // Poll status every 10s until processed
  for (let i = 0; i < 30; i++) {
    await new Promise((r) => setTimeout(r, 10000));
    const t = await getToken();
    const s = await fetchRaw(
      "https://marketplace.walmartapis.com/v3/feeds/" + feedId + "?includeDetails=true",
      {
        headers: {
          Accept: "application/json",
          "WM_SEC.ACCESS_TOKEN": t,
          "WM_SVC.NAME": "Walmart Marketplace",
          "WM_QOS.CORRELATION_ID": uuid(),
        },
      }
    );
    const d = JSON.parse(s.body.toString());
    console.log(
      "[" + i + "] " + d.feedStatus + " - processed: " + (d.itemsReceived || 0) + "/" + (d.itemsSucceeded || 0) + " success, " + (d.itemsFailed || 0) + " failed"
    );
    if (d.feedStatus === "PROCESSED" || d.feedStatus === "ERROR") {
      fs.writeFileSync("data/feed_result.json", JSON.stringify(d, null, 2));
      console.log("Saved data/feed_result.json");
      return;
    }
  }
}

submit().catch((e) => {
  console.error(e);
  process.exit(1);
});
