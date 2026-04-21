#!/usr/bin/env node

const { latexToOMML } = require("latex-to-omml");

async function main() {
  let input = "";
  process.stdin.setEncoding("utf8");

  for await (const chunk of process.stdin) {
    input += chunk;
  }

  const payload = JSON.parse(input || "{}");
  const items = Array.isArray(payload.items) ? payload.items : [];
  const results = [];

  for (const item of items) {
    const id = item && Object.prototype.hasOwnProperty.call(item, "id") ? item.id : null;
    const latex = typeof item?.latex === "string" ? item.latex : "";
    const displayMode = Boolean(item?.displayMode);

    try {
      const omml = await latexToOMML(latex, { displayMode });
      results.push({ id, ok: true, omml });
    } catch (error) {
      results.push({
        id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  process.stdout.write(JSON.stringify({ results }));
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exit(1);
});
