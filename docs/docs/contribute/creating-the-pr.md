# Submitting your Local Changes

> If you're new to contributing to open-source repositories using GitHub or find the instructions on this page confusing, we recommend signing up for one of our [Sharing is Caring events](https://pnp.github.io/sharing-is-caring/#pnp-sic-events). These events are completely free and designed to guide you through the process.

Once you have everything created and your code/script sample is fully functional, you can proceed with creating a pull request (PR) to the CLI for Microsoft 365 repository to include your changes in the next release. Before submitting the PR, ensure that you have tested your code locally and that your tests provide 100% code coverage.

!!! important

    All git commands should be executed from the root of your locally cloned fork.

If this is your first PR, make sure you have created a link to the upstream repository of the [CLI for Microsoft 365](https://github.com/pnp/cli-microsoft365).

```bash
# Check if you have a remote pointing to the CLI for Microsoft 365 repo:
git remote -v

# If you see a pair of remotes (fetch & pull) that point to https://github.com/pnp/cli-microsoft365
# Then you are ok. Otherwise, you need to add one

# Add a new remote named "upstream" and point it to the CLI repo
git remote add upstream https://github.com/pnp/cli-microsoft365.git
```

## Check in your Local Changes

Before proceeding, ensure that all your files are committed to your branch. If you're using Visual Studio Code, you can achieve this by selecting `Source Control` from the activity bar and providing a commit name in the `Source Control` tab. Generally, the commit name should be the title of the issue you are working on. Then click the commit button to safely commit your files to your local branch.

If you're solely working with git, follow these steps:

```bash
# Stash all your files
git add .

# Commit your stash
git commit -m 'Adds command spo group get'
```

## Rebase the Latest Changes

Next up, we want to make sure our branch is up to date with the latest changes from the upstream branch. 

```bash
# Fetch all the change from upstream
git fetch upstream

# Rebase them into your branch
git pull --rebase upstream main
```

## Push your Local Branch

Now we need to get all our local changes in our forked repository. This way we can create a new PR using GitHub.
```bash
# Push your changes to your fork
git push origin

# If this fails due to some failed refs, then you can use the --force flag
git push origin --force
```

## Creating the Pull Request

With all your changes pushed to your forked repository, you can navigate to the CLI for Microsoft 365 repository and go to the [Pull Requests](https://github.com/pnp/cli-microsoft365/pulls) tab. Here, you'll find a green "New pull request" button that allows you to create a PR based on the branch you just pushed. Click on `Compare & pull request` to start creating your PR. Make sure to provide a descriptive title for the PR, such as 'Add command spo group get'. In the description, explain what you have accomplished—whether it's a new command, bug fix, or a minor update in the documentation. The more information you provide, the quicker your PR can be reviewed and merged. Finally, mention the issue that this PR is based on and include it in the following format: `Closes #[Id of the issue]`.

Once you have filled in all the necessary details, click on the `Create pull request` button and await approval from one of our maintainers.

## What Happens Next?

Once your hard work is submitted then it's up to us. One of the CLI for Microsoft 365 maintainers will review your PR and provide feedback if necessary. If everything looks good, your PR will be merged into the main branch and will be included in the next release of the CLI for Microsoft 365. To get some more insight into what happens during this process, check out the [What to expect during a Pull request review](./expect-during-PR.md).
