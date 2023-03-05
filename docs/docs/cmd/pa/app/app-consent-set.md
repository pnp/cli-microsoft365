# pa app consnent set

Makes sure users can bypass the API Consent window for the selected canvas app

## Usage

```sh
m365 pa app consent set [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-n, --name <name>`
: The name of the Microsoft Power App that should bypass the API consent

`-e, --enabled <enabled>`
: Set to true to enable the Microsoft App to bypass the API consent, or false to disable it. Valid values are `true` or `false`

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Remarks

This command only works for canvas apps

## Examples

Enabled the bypass consent for the specified canvas app

```sh
m365 pa app consent set --environment 4be50206-9576-4237-8b17-38d8aadfaa36 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --enabled
```

Disables the bypass consent for the specified canvas app

```sh
m365 pa app consent set --environment 4be50206-9576-4237-8b17-38d8aadfaa36 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --enabled false --confirm
```

## Response

The command won't return a response on success.
