# teams funsettings list

Lists fun settings for the specified Microsoft Teams team

## Usage

```sh
m365 teams funsettings list [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the team for which to list fun settings

--8<-- "docs/cmd/_global.md"

## Examples

List fun settings of a Microsoft Teams team

```sh
m365 teams funsettings list --teamId 83cece1e-938d-44a1-8b86-918cf6151957
```

## Response

=== "JSON"

    ```json
    {
      "allowGiphy": true,
      "giphyContentRating": "moderate",
      "allowStickersAndMemes": true,
      "allowCustomMemes": true
    }
    ```

=== "Text"

    ```text
    allowCustomMemes     : true
    allowGiphy           : true
    allowStickersAndMemes: true
    giphyContentRating   : moderate
    ```

=== "CSV"

    ```csv
    allowGiphy,giphyContentRating,allowStickersAndMemes,allowCustomMemes
    1,moderate,1,1
    ```

==="Markdown"

    ```md
  # teams funsettings list --teamId "83cece1e-938d-44a1-8b86-918cf6151957"

Date: 5/7/2023

Property | Value
---------|-------
allowGiphy | true
giphyContentRating | moderate
allowStickersAndMemes | true
allowCustomMemes | true
     ```
