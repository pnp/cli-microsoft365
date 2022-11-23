# Creating a new command

Awesome! Good ideas are invaluable for every product. 

!!! note

    Before you start hacking away, please check if there is no similar idea already listed in the [issue list](https://github.com/pnp/cli-microsoft365/issues). If not, please create a new issue describing your idea.

Once we agree on the feature scope and architecture, the feature will be ready for building. Don't hesitate to mention this in the issue if you'd like to build the feature yourself. When you start creating a new command you will need to write the command logic, unit tests, and documentation. Here you can find a detailed guide on how to create a new command from scratch.

## Creating a new branch

Once you have cloned the repository, create a new branch to work on using the command `git checkout -b [branch name] main`. This branch will contain your changes and will be used to create the pull request.

## Step-by-step guide

We will guide you through a workflow on how to create a new command from scratch, starting from an example issue. This example will be used throughout the step-by-step guide to provide more insight into a realistic scenario. 

---

## New command: Get site group

<h4> Usage </h4>

m365 spo group get [options]

<h4> Description </h4>

Gets site group

<h4> Options </h4>

Option | Description
-- | --
`-u, --webUrl <webUrl>` | URL of the site where the group is located.
`-i, --id  [id]` | ID of the site group to get. Use either `id`, `name` or `associatedGroup` but not multiple.
`--name  [name]` | ID of the site group to get. Use either `id`, `name` or `associatedGroup` but not multiple.
`--associatedGroup [associatedGroup]` | ID of the site group to get. Available values: `Owner`, `Member`, `Visitor`. Use either `id`, `name` or `associatedGroup` but not multiple.

<h4> Examples </h4>

Get group with ID 7 for web https://contoso.sharepoint.com/sites/project-x

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --id 7
```

Get group with name Team Site Members for web https://contoso.sharepoint.com/sites/project-x

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --name "Team Site Members"
```

Get the associated owner group of a specified site

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --associatedGroup Owner
```

---
