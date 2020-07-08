# Minimal Path to Awesome

The shortest way to prepare your local copy of the project for development and testing.

## Install prerequisites

Before you start contributing to this project, you will need Node.js `>= 12.0.0` installed. This project has been tested with the LTS version of Node.js and the version of NPM that comes with it.

## Get the local version of the CLI

- fork this repository
- clone your fork
- in the command line:
  - run `npm i` to restore dependencies
  - run `npm run build` to build the project
  - run `npm test` to run test and check current code coverage
  - run `npm link` to install the project locally. This is useful if you want to test your changes to the CLI in the CLI itself. After linking the local package, you can start your local version of the CLI by typing in the command line `m365`.

> If you installed the CLI globally using the `npm i -g @pnp/cli-microsoft365` command, we recommend that you uninstall it first, before running `npm link`

After changing the code, run the `npm run build` command to rebuild the project and see your changes integrated in the local version of the CLI.

If you renamed files:

- in the command line:
  - run `npm run clean` to clean up the output folder
  - run `npm run build` to rebuild the project
  - run `npm link` to reinstall the project locally. Without this step, you will get an error, when trying to start the local version of the CLI.

### Documentation

CLI for Microsoft 365 uses [MkDocs](http://www.mkdocs.org) to publish documentation pages. See more information about installing MkDocs on your operating system at [http://www.mkdocs.org/#installation](http://www.mkdocs.org/#installation).

CLI for Microsoft 365 documentation uses the `mkdocs-material` theme. See more information about installing mkdocs-material on your operating system at [https://squidfunk.github.io/mkdocs-material](https://squidfunk.github.io/mkdocs-material).

Once you have MkDocs installed on your machine, in the command line:

- run `cd ./docs/manual` to change directory to where the manual pages are stored
- run `mkdocs serve` to start the local web server with MkDocs and view the documentation in the web browser

Alternatively, you can use the mkdocs-material Docker image to test documentation. In order to use docker you are required to specify the correct material version:

- on macOS:
  - run `cd ./docs/manual` to change directory to where the manual pages are stored
  - run `docker run --rm -it -p 8000:8000 -v ${PWD}:/docs squidfunk/mkdocs-material:3.1.0` to start the local web server with MkDocs and view the documentation in the web browser
- on Windows:
  - run `docker run --rm -it -p 8000:8000 -v c:/projects/cli-microsoft365/docs/manual:/docs squidfunk/mkdocs-material:3.1.0` to start the local web server with MkDocs and view the documentation in the web browser
