> Thank you for submitting your change to the project, we really appreciate your help.
>
> For us to be able to review and merge your changes as quickly as possible, please review our pull request process.
>
> _(DELETE THIS PARAGRAPH AFTER READING)_
>

> ### Pull Request Title
>
> Please ensure that you have included a meaningful title describing your change.
>
> For example...
>
> - Adds 'spo orgassetslibrary remove' command. Closes #1042
> - Extends 'spo site list' command with support for returning deleted sites. Closes #1335
> - Fixes 'spo listitem add' command. Closes #1297
>
> _(DELETE THIS SECTION AFTER READING)_
>

> ### Linked Issue
>
> Please ensure that your change is related to an open issue, we will not accept changes without a related issue.
>
> You should link your pull request to an open issue using the `Closes` keyword, referencing the issue using `#` and the issue id. This will ensure that the issue is automatically closed when your pull request is merged.
>
> _(DELETE THIS SECTION AFTER READING)_
>

Closes #

> ### One PR = One Commit
>
> We prefer to review pull requests that contain a single commit as this makes it easier to keep the history clear for everyone and track what's changed.
>
> If you are new to git, we'll squash your commits for you when merging your changes.
>
> To squash your commits, use the interactive rebase method, `git rebase -i HEAD~<commits>`, squashing the commits into a single commit and changing your commit message to be the same as your pull request title.
>
> _(DELETE THIS SECTION AFTER READING)_
>

> ### New Command = Five Files
>
> If your pull request contains the changes for a new command, it should contain five changed files.
>
> For example, a pull request for adding 'teams tab add' command, would include...
>
> - src/m365/teams/commands.ts
> - src/m365/teams/commands/tab/teams-tab-add.ts
> - src/m365/teams/commands/tab/teams-tab-add.spec.ts
> - docs/manual/mkdocs.yml
> - docs/manual/docs/cmd/teams/tab/tab-add.md
>
> _(DELETE THIS SECTION AFTER READING)_
>

> ### Master Branch Only
>
> You should only submit your changes to the **master** branch.
>
> Pull requests submitted to other branches will be rejected.
>
> _(DELETE THIS SECTION AFTER READING)_
>

> ### Merge Conflicts
>
> Merge conflicts occur when the branch you want to merge is out of date. To ensure this does not happen, you should ensure that your branch is up-to-date before submitting your pull request by executing the below git commands.
>
> `git checkout master`
>
> `git pull upstream master`
>
> `git push origin master`
>
> `git checkout <your-branch>`
>
> `git rebase master`
>
> `git push origin <your-branch> -f`
>
> _(DELETE THIS SECTION AFTER READING)_
>
