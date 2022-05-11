# planner task reference add

Adds a new reference to a Planner task.

## Usage

```sh
m365 planner task reference add [options]
```

## Options

`-i, --taskId <taskId>`
: ID of the task.

`-u, --url <url>`
: URL location of the reference.

`--alias [alias]`
: A name alias to describe the reference.

`--type [type]`
: Used to describe the type of the reference. Types include: `PowerPoint`, `Word`, `Excel`, `Other`.

--8<-- "docs/cmd/_global.md"

## Examples

Add a new reference with the url _https://www.microsoft.com_ to a Planner task with the id _2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2_

```sh
m365 planner task reference add --taskId "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" --url "https://www.microsoft.com"
```

Add a new reference with the url _https://www.microsoft.com_ and with the alias _Parker_ to a Planner task with the id _2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2_

```sh
m365 planner task reference add --taskId "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" --url "https://www.microsoft.com" --alias "Parker"
```

Add a new reference with the url _https://www.microsoft.com_ and with the type Excel to a Planner task with the id _2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2_

```sh
m365 planner task reference add --taskId "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" --url "https://www.microsoft.com" --type "Excel"
```