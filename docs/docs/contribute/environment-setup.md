# Setup your development environment

The shortest way to prepare your local copy of the project for development and testing.

!!! info "Install prerequisites"

    Before you start contributing to this project, you will need **Node.js@16** and **npm@8** installed.

## Get the local version of the CLI

### Setting up your GitHub fork

1. Navigate to the [CLI for Microsoft 365](https://github.com/pnp/cli-microsoft365) repository
1. Fork this repository. On the GitHub repository page, click the "Fork" button to create your own copy of the repository.
    - Make sure the option 'Copy the `main` branch only' is selected 
1. Clone the forked repository to your local machine. Clone the repository using Git to your local machine using the command `git clone [repository URL]`

> More insights about forking a repository can be found here: [GitHub Docs - Fork a repo](https://docs.github.com/en/get-started/quickstart/fork-a-repo#forking-a-repository)

### Setting up your local project

After you've cloned your fork onto your device, you can navigate to the project directory and start executing the following commands to get the project running.

1. `npm i`: restore all dependencies of the project
1. `npm run build`: build the entire project
1. `npm link`: create a link/reference to your local project. This allows you to reference your locally installed CLI instead of the npm-hosted package.

That's it! If you now run `m365 version` you will see that you are now using the beta version of CLI for Microsoft 365 in your shell!

!!! tip

    If you installed the CLI globally using the `npm i -g @pnp/cli-microsoft365` command, we recommend that you uninstall it first, before running `npm link`

## Visual Studio Code extensions

It doesn't matter which IDE you are using although we recommend using Visual Studio Code.
When using VS Code, the following extensions will come in handy:

Extension | Why is it useful?
--- | ---
[Mocha Test Explorer](https://marketplace.visualstudio.com/items?itemName=hbenl.vscode-mocha-test-adapter) | This extension will help you when writing tests. It is capable of running individual tests and debugging. You will also get a nice overview of all tests within the project.
[ESLint](https://marketplace.visualstudio.com/items?itemName=dbaeumer.vscode-eslint) | We use ESLint to monitor consistency within the project. By installing this extension, you are notified of problems while writing code.

## Npm scripts

During your development, you will need a set of commands to test, run, and validate your code. CLI for Microsoft 365 comes with a set of commands that you can use for each occasion.

Command | Description
------ | ---
`npm run watch` | Builds the entire project first. After this, a watcher will make sure that every time a file is saved, an incremental build is triggered. This means that not the entire project is rebuilt but only the changed files. **You typically use this command while developing**.
`npm run build` | Builds the entire project.
`npm run clean` | Clean the output directory. All built files will be deleted.
`npm run test` | Run all tests, check all ESLint rules, ... This is a combination of `npm run test:cov` and `npm run lint`. This is what happens in our GitHub workflows when creating a PR.
`npm run test:cov` | Run all tests and create a coverage report.
`npm run test:test` | Run all tests.
`npm run lint` | Run all ESLint rules.
