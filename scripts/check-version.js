const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "18") {
  console.error("Node version must be 18");
  process.exitCode = 1;
}
