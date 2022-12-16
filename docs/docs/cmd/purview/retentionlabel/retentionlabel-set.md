# purview retentionlabel set

Update a retention label

## Usage

```sh
m365 purview retentionlabel set [options]
```

## Options

`-i, --id <id>`
: The Id of the retention label.

`--behaviorDuringRetentionPeriod [behaviorDuringRetentionPeriod]`
: Specifies how the behavior of a document with this label should be during the retention period. Allowed values: `doNotRetain`, `retain`, `retainAsRecord`, `retainAsRegulatoryRecord`.

`--actionAfterRetentionPeriod [actionAfterRetentionPeriod]`
: Specifies the action to take on a document with this label applied after the retention period. Allowed values: `none`, `delete`, `startDispositionReview`.

`--retentionDuration [retentionDuration]`
: The number of days to retain the content.

`-t, --retentionTrigger [retentionTrigger]`
: Specifies whether the retention duration is calculated from the content creation date, labeled date, or last modification date. Allowed values: `dateLabeled`, `dateCreated`, `dateModified`, `dateOfEvent`.

`--defaultRecordBehavior [defaultRecordBehavior]`
: Specifies the locked or unlocked state of a record label when it is created. Allowed values: `startLocked`, `startUnlocked`.

`--descriptionForUsers [descriptionForUsers]`
: The label information for the user.

`--descriptionForAdmins [descriptionForAdmins]`
: The label information for the admin.

`--labelToBeApplied [labelToBeApplied]`
: Specifies the replacement label to be applied automatically after the retention period of the current label ends.

--8<-- "docs/cmd/_global.md"

## Examples

Update a retention label so that it retains documents as records and deletes them after one year.

```sh
m365 purview retentionlabel set --id c37d695e-d581-4ae9-82a0-9364eba4291e --behaviorDuringRetentionPeriod retainAsRecord --actionAfterRetentionPeriod delete --retentionDuration 365
```

Update a retention label so that it retains documents as regulatory records and starts a disposition review one year after the last modification date.

```sh
m365 purview retentionlabel set --id c37d695e-d581-4ae9-82a0-9364eba4291e --behaviorDuringRetentionPeriod retainAsRegulatoryRecord --actionAfterRetentionPeriod startDispositionReview --retentionDuration 365 --retentionTrigger dateModified
```

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

## Response

The command won't return a response on success.
