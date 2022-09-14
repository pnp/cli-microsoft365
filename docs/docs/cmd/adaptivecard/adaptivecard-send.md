# adaptivecard send

Sends adaptive card to the specified URL

## Usage

```sh
m365 adaptivecard send [options]
```

## Options

`-u, --url <url>`
: URL where to send the card to

`-t, --title [title]`
: Title of the card

`-d, --description [description]`
: Contents of the card

`-i, --imageUrl [imageUrl]`
: URL of the image to include on the card

`-a, --actionUrl [actionUrl]`
: URL that users should be sent to after clicking the **View** button on the card

`--card [card]`
: Card definition

`--cardData [cardData]`
: Card data. If your card is a card template, using cardData you can apply data to it. If you specify cardData, unknown options will be ignored.

--8<-- "docs/cmd/_global.md"

## Remarks

Using this command you can send either a predefined or a custom adaptive card to the specified URL. To send a predefined adaptive card, specify one or more options: `title`, `description`, `imageUrl`, `actionUrl`. To specify a custom card, specify the card's JSON contents using the `card` option.

When sending both predefined and custom cards, you can specify arbitrary options. With predefined cards, these options will be listed in a FactSet. With custom card, you control how the data will be presented on the card.

The predefined card is automatically adjusted based on which options have been specified (`title`, `description`, `imageUrl`, `actionUrl`). If you don't specify a particular option, that portion of the card will not be included in the sent card.

If your custom card is a card template (card with placeholders like `${title}`), you can fill it with data either by specifying the complete data object using the `cardData` option, or by passing any number of arbitrary options that will be mapped onto the card. The arbitrary properties should not match any of the global options like `output`, `query`, `debug`, etc. Data options like `title`, `description`, `imageUrl` and `actionUrl` will be mapped onto the card as well.

## Examples

Send a predefined adaptive card with just title

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --title "CLI for Microsoft 365 v3.4"
```

Send a predefined adaptive card with just description

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --description "New release of CLI for Microsoft 365"
```

Send card with title and description

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --title "CLI for Microsoft 365 v3.4" --description "New release of CLI for Microsoft 365"
```

Send card with title, description and image

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --title "CLI for Microsoft 365 v3.4" --description "New release of CLI for Microsoft 365" --imageUrl "https://contoso.com/image.gif"
```

Send card with title, description and action

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --title "CLI for Microsoft 365 v3.4" --description "New release of CLI for Microsoft 365" --actionUrl "https://aka.ms/cli-m365"
```

Send card with title, description, image and action

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --title "CLI for Microsoft 365 v3.4" --description "New release of CLI for Microsoft 365" --imageUrl "https://contoso.com/image.gif" --actionUrl "https://aka.ms/cli-m365"
```

Send card with title, description, action and unknown options

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --title "CLI for Microsoft 365 v3.4" --description "New release of CLI for Microsoft 365" --actionUrl "https://aka.ms/cli-m365" --Version "v3.4.0" --ReleaseNotes "https://pnp.github.io/cli-microsoft365/about/release-notes/#v340"
```

Send custom card without any data

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --card '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"CLI for Microsoft 365 v3.4"},{"type":"TextBlock","text":"New release of CLI for Microsoft 365","wrap":true}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"https://aka.ms/cli-m365"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}'
```

Send custom card with just title merged

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --card '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}' --title "CLI for Microsoft 365 v3.4"
```

Send custom card with all known options merged

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --card '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${actionUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}' --title "CLI for Microsoft 365 v3.4" --description "New release of CLI for Microsoft 365" --imageUrl "https://contoso.com/image.gif" --actionUrl "https://aka.ms/cli-m365"
```

Send custom card with unknown option merged

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --card '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${Title}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}' --Title "CLI for Microsoft 365 v3.4"
```

Send custom card with card data

```sh
m365 adaptivecard send --url https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547 --card '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}' --cardData '{"title":"Publish Adaptive Card Schema","description":"Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation.","creator":{"name":"Matt Hidinger","profileImage":"https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg"},"createdUtc":"2017-02-14T06:08:39Z","viewUrl":"https://adaptivecards.io","properties":[{"key":"Board","value":"Adaptive Cards"},{"key":"List","value":"Backlog"},{"key":"Assigned to","value":"Matt Hidinger"},{"key":"Due date","value":"Not set"}]}'
```
