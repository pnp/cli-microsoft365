# docs

Returns the CLI for Microsoft 365 docs webpage URL

## Usage

```sh
m365 docs [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

If config setting `autoOpenLinksInBrowser` is configured to true, the command will automatically open the CLI for Microsoft 365 docs webpage in the default browser. [cli config set](../cmd/cli/config/config-set.md)

## Examples

Returns the CLI for Microsoft 365 docs webpage URL

```sh
m365 docs
```

## Response

=== "JSON"

    ```json
    "https://pnp.github.io/cli-microsoft365/"
    "Use a web browser to open the CLI for Microsoft 365 docs webpage URL"
    ```

=== "Text"

    ```text
    https://pnp.github.io/cli-microsoft365/
    Use a web browser to open the CLI for Microsoft 365 docs webpage URL
    ```

=== "CSV"

    ```csv
    https://pnp.github.io/cli-microsoft365/
    Use a web browser to open the CLI for Microsoft 365 docs webpage URL
    ```

=== "Markdown"

    ```md
    https://pnp.github.io/cli-microsoft365/
    Use a web browser to open the CLI for Microsoft 365 docs webpage URL
    ```
