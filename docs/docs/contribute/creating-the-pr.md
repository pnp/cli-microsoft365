# Submitting the new command

> If you aren't familiar with how to contribute to open-source repositories using GitHub, or if you find the instructions on this page confusing, [sign up](https://forms.office.com/Pages/ResponsePage.aspx?id=KtIy2vgLW0SOgZbwvQuRaXDXyCl9DkBHq4A2OG7uLpdUREZVRDVYUUJLT1VNRDM4SjhGMlpUNzBORy4u) for one of our [Sharing is Caring](https://pnp.github.io/sharing-is-caring/#pnp-sic-events) events. It's completely free, and we'll guide you through the process.

With everything created and your code/script sample fully functional, we can create a PR to the CLI for Microsoft 365 repository to include it in the next release. Before submitting the PR, make sure that you tested your code locally and that your tests are at 100% code coverage. 

!!! important

    Every git command should be executed from the root of your locally cloned fork

If this is your first PR, make sure you have created a link to the upstream repository of the [CLI for Microsoft 365](https://github.com/pnp/cli-microsoft365).

```bash
# Check if you have a remote pointing to the Microsoft repo:
git remote -v

# If you see a pair of remotes (fetch & pull) that point to https://github.com/pnp/cli-microsoft365
# Then you are ok. Otherwise, you need to add one

# Add a new remote named "upstream" and point it to the CLI repo
git remote add upstream https://github.com/pnp/cli-microsoft365.git
```

## Check in your local changes

Before we continue, we should make sure that all our files are committed in our branch. This can be achieved via git commands or, if you are using Visual Studio Code, from the activity bar, selecting `Source Control`. On the tab `Source Control`, you can give your commit a name. Generally, this will be the title of the issue you are working on. Then you press the button commit and your files should be safely committed to your local branch.

If you are solely working with git, you will need to do the following.

```bash
# Stash all your files
git add .

# Commit your stash
git commit -m 'Adds command spo group get'
```

## Rebase the latest changes

Next up, we want to make sure our branch is up to date with the latest changes from the upstream branch. 

```bash
# Fetch all the change from upstream
git fetch upstream

# Rebase them into your branch
git pull --rebase upstream main
```

## Push your local branch

Now we need to get all our local changes in our forked repository. This way we can create a new PR using GitHub.
```bash
# Push your changes to your fork
git push origin

# If this fails due to some failed refs, then you can use the --force flag
git push origin --force
```

## Creating the Pull Request

With everything pushed to your forked repository, you can navigate to the CLI for Microsoft 365 repository and then to the tab [Pull Requests](https://github.com/pnp/cli-microsoft365/pulls). Here you will get a green popup to create a PR based on the branch you just published. Click on `Compare & pull request` to start the creation of your PR. Be sure to give the PR a descriptive title e.g. 'Add command spo group get'. Then, explain in the description what you have made. Is it a new command, a bug fix, or a minor update in the docs? The clearer the information you provide, the quicker your PR can be verified and merged. Finally, in the description, be sure to mention the issue that this PR is based on and included it in the following form `Closes #[Id of the issue]`.

With all this ready, you can click on the button `Create pull request` and then you can await approval from one of our maintainers. 
