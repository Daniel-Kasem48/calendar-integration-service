import { createApp } from "./app";
async function main() {
  const app = await createApp();
  app.enableShutdownHooks();
  const port = Number(process.env.PORT || 3000);
  if (!Number.isInteger(port) || port < 1 || port > 65535)
    throw Error("PORT must be between 1 and 65535");
  await app.listen(port, process.env.HOST || "127.0.0.1");
  console.log(`Calendar GraphQL service listening on port ${port}`);
}
main().catch(() => {
  console.error(
    "Startup failed: check tenant configuration, PUBLIC_URL and PORT.",
  );
  process.exitCode = 1;
});
