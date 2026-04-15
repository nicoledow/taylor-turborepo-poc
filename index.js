import express from "express";

const app = express();
const port = process.env.PORT ?? 3000;

app.get("/hello", (req, res) => {
  res.json({ message: "hello world" });
});

app.listen(port, () => {
  console.log(`Listening on http://localhost:${port}`);
});
