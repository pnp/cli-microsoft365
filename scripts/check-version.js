const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "20") {
  console.error("Node version must be 20");
  process.exitCode = 1;
}
