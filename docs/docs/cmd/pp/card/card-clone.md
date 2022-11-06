# pp card clone

Clones a specific Microsoft Power Platform card in the specified Power Platform environment

## Usage

```sh
pp card clone [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`--newName <newName>`
: The name of the new card.

`-i, --id [id]`
: The id of the card. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The name of the card. Specify either `id` or `name` but not both.

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

Clones a specific card in a specific environment based on name.

```sh
m365 pp card clone --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "CLI 365 Card" --newName "CLI 365 new Card"
```

Clones a specific card in a specific environment based on name as admin.

```sh
m365 pp card clone --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "CLI 365 Card" --newName "CLI 365 new Card" --asAdmin 
```

Clones a specific card in a specific environment based on id.

```sh
m365 pp card clone --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "408e3f42-4c9e-4c93-8aaf-3cbdea9179aa" --newName "CLI 365 new Card"
```

Clones a specific card in a specific environment based on id as admin.

```sh
m365 pp card clone --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "408e3f42-4c9e-4c93-8aaf-3cbdea9179aa" --newName "CLI 365 new Card" --asAdmin
```

## Response

=== "JSON"

    ```json
    {
      "CardIdClone": "80cff342-ddf1-4633-aec1-6d3d131b29e0"
    }
    ```

=== "Text"

    ```text
    CardIdClone: 80cff342-ddf1-4633-aec1-6d3d131b29e0
    ```

=== "CSV"

    ```csv
    CardIdClone
    80cff342-ddf1-4633-aec1-6d3d131b29e0
    ```
