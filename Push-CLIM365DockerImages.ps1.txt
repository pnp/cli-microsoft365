# This script iterates the array items, building a new Docker image for each version specified and pushes the image to the Docker Hub
# Script should be executed in the same folder as the Dockerfile
# Requires you to be logged into Docker CLI with an account that has push rights to m365pnp organisation
@("3.6.0-beta.b83cb79","next") | ForEach-Object { $_; docker build --pull --no-cache --rm -t m365pnp/cli-microsoft365:$_ "." --build-arg CLI_VERSION=$_; docker push m365pnp/cli-microsoft365:$_ }