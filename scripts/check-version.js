const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "26") {
  console.error("Node version must be 26");
  process.exitCode = 1;
}
