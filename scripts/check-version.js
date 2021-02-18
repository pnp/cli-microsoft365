const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "14") {
  console.error("Node version must be 14");
  process.exitCode = 1;
}
