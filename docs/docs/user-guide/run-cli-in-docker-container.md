# Run CLI for Microsoft 365 in a Docker Container

You can use Docker to run a standalone Linux container with CLI for Microsoft 365 and PowerShell pre-installed, with command completion (tab) automatically configured for you in both bash and PowerShell, without having to install any of the required dependencies on your host machine.

## Prerequisites

To use the published Docker images, you will first need to have Docker installed and configured on your host machine. Please refer to the Docker documentation on how to [install Docker](https://docs.docker.com/get-docker/) on Windows, Mac and Linux.

## Install and run latest

To install and run the latest version of CLI for Microsoft 365, use the `docker run` command and specify the CLI docker image name.

```
docker run --rm -it m365pnp/cli-microsoft365:latest
```

Executing this command for the first time will download the image onto your machine and start the container and invoke an interactive session automatically `(-it)`, displaying a bash shell prompt.

You can exit from the prompt by closing the terminal window or typing `exit`. When you exit the container it will be automatically stopped to free up system resources `(--rm)`.


Alternatively, you can use PowerShell as the default shell by passing `pwsh` into the `docker run` command after the image name.

```
docker run --rm -it m365pnp/cli-microsoft365:latest pwsh
```

!!! info
    Authentication information is not persisted in the Docker container. When you exit from the container, you will need to authenticate with your Microsoft 365 tenant the next time you run the container.

## Install and run beta

We regularly release beta versions of the CLI, to install and run the latest beta release use the `next` tag.

```
docker run --rm -it m365pnp/cli-microsoft365:next
```

## Install and run specific versions

We have published Docker images for every minor release of v3 of CLI for Microsoft 365 to date. 

To install and run a specific version of the CLI, state the version number as a tag after the image name.

```
docker run --rm -it m365pnp/cli-microsoft365:3.0.0
```

You can also install and run specific beta versions of the CLI, state the beta version as a tag after the image name.

```
docker run --rm -it m365pnp/cli-microsoft365:3.4.0-beta.0dbd08d
```

## Execute script in container

In scenarios where you may already have a script that uses the CLI for Microsoft 365 and you want to execute it within the container, you can use a volume mount to share files on your host machine with the Docker container.

For example, lets say we have a script called `test.sh` and we want to execute that script inside the container. We can do this by mapping the current working directory on our host machine to the working directory in the container `(-v)`, pass `bash` as the shell we want to use and the name of the file that we want to execute as additional arguments.

```
docker run -it -v ${PWD}:/home/cli-microsoft365/scripts m365pnp/cli-microsoft365:latest bash scripts/test.sh
```

Alternatively, if we want to execute a PowerShell script, you can do this in the same way.

```
docker run -it -v ${PWD}:/home/cli-microsoft365/scripts m365pnp/cli-microsoft365:latest pwsh scripts/test.ps1
```

!!! info
    We have created a non-root user called `cli-microsoft365` inside the container.  When the container starts, the working directory is set to the home directory of this user, hence the need to add `/home/cli-microsoft365` to the volume mapping.

## Set Environment Variables

In scenarios where you need to set environment variables, for example, you want to use a custom Azure AD identity when logging into your Microsoft 365 tenant using the CLI. You can set these variables by passing them in as options arguments `(-e)` into the `docker run` command.

```sh
docker run --rm -it -e "CLIMICROSOFT365_AADAPPID=51078274-0353-4f6a-b9f5-8674ab2e524c" -e "CLIMICROSOFT365_TENANT=9455bc83-d5af-4ccf-93f6-0af3f71aaf8e" m365pnp/cli-microsoft365:latest
```

## Combining script and environment variables

Combining scripts and environment variables is a powerful way to run the CLI in Docker, we can set environment variables which we can reference in the script that is executed in the running container and also.

```sh
docker run --rm -it -v ${PWD}:/home/cli-microsoft365/scripts -e "CLIMICROSOFT365_AADAPPID=da049853-dd90-49df-aa21-4e0c8b646a36" -e "CLIMICROSOFT365_TENANT=e8954f17-a373-4b61-b54d-45c038fe3188" -e "M365_USER=user@contoso.com" -e "M365_PASSWORD=password" m365pnp/cli-microsoft365:next pwsh scripts/script.ps1
```

We can reference the environment variables passed in to the `docker run` command and use them in the script, in this example, passing the username and password variables into the `m365 login` command to login in to Microsoft 365 using password authentication.

```
m365 login --authType password --userName $env:M365_USER --password $env:M365_PASSWORD
```

## Update Docker Image

We will be regularly updating the images of the `latest` and `next` tags, to ensure you have the most upto date version of these images, you can update your local image using `docker pull` specifying the version you want to update using the relevant tag.

```
docker pull m365pnp/cli-microsoft365:latest
```

## Uninstall Docker Image

If you would like to remove an image from your host machine, you can use the `rmi` command, specifying the version you wish to remove as a tag after the image name.

```
docker rmi m365pnp/cli-microsoft365:latest
```

## Published Tags

See the list of available tags on the m365pnp/cli-microsoft365 repository on [Docker Hub](https://hub.docker.com/repository/docker/m365pnp/cli-microsoft365/).
