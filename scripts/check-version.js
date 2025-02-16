const nodeVersion = process.versions.node.split('.')[0];

if (nodeVersion !== "22") {
  console.error("Node version must be 22");
  process.exitCode = 1;
}
