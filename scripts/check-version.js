const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "16") {
  console.error("Node version must be 16");
  process.exitCode = 1;
}
