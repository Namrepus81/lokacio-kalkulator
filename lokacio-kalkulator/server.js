const http = require("http");
const fs = require("fs");
const path = require("path");

const port = 8765;
const root = __dirname;

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
};

const server = http.createServer((request, response) => {
  const requestUrl = new URL(request.url, `http://localhost:${port}`);
  const safePath = path.normalize(decodeURIComponent(requestUrl.pathname)).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(root, safePath === "\\" || safePath === "/" ? "index.html" : safePath);

  if (!filePath.startsWith(root)) {
    response.writeHead(403);
    response.end("Tiltott utvonal");
    return;
  }

  fs.readFile(filePath, (error, content) => {
    if (error) {
      response.writeHead(404);
      response.end("Nincs ilyen fajl");
      return;
    }

    response.writeHead(200, { "Content-Type": contentTypes[path.extname(filePath)] || "application/octet-stream" });
    response.end(content);
  });
});

server.listen(port, () => {
  console.log(`Lokacio Kalkulator: http://127.0.0.1:${port}`);
});
