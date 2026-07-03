import express from "express";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const app = express();
const __dirname = dirname(fileURLToPath(import.meta.url));
const dist = join(__dirname, "dist");

app.use(express.static(dist));
app.get("*", (_req, res) => res.sendFile(join(dist, "index.html")));

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Listening on :${port}`));
