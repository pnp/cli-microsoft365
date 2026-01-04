const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "24") {
  console.error("Node version must be 24");
  process.exitCode = 1;
}
