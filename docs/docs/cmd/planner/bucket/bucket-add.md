# planner bucket add

Adds a new Microsoft Planner bucket

## Usage

```sh
m365 planner bucket add [options]
```

## Options

`-n, --name <name>`
: Name of the bucket to add.

`--planId [planId]`
: ID of the plan to which the bucket belongs. Specify either `planId` or `planTitle` but not both.

`--planTitle [planTitle]`
: Title of the plan to which the bucket belongs. Specify either `planId` or `planTitle` but not both.

`--planName [planName]`
: (deprecated. Use `planTitle` instead) Title of the plan to which the bucket belongs.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--orderHint [orderHint]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/en-us/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

--8<-- "docs/cmd/_global.md"

## Examples

### Add a bucket with a name and an order hint.

Adds a new Microsoft Planner bucket with a name and an order hint. This will be created in a Planner plan based on its id.


``` sh title="Command"
m365 planner bucket add --name "My Planner Bucket" --planId "xqQg5FS2LkCp935s-FIFm2QAFkHM" --orderHint " !"
```

``` json title="Output"
{
  "name": "My Planner Bucket",
  "planId": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
  "orderHint": "8585363889524958496",
  "id": "ttEB_Uj690STdR3GC1MIDZgANq1U"
}
```

### Add a bucket based on plan title and group name.

Adds a new Microsoft Planner bucket with a name This will be created in a Planner plan based on its title and the plans group name.


``` sh title="Command"
m365 planner bucket add --name "My Planner Bucket" --planTitle "My Planner Plan" --ownerGroupName "My Planner Group"
```

``` json title="Output"
{
  "name": "My Planner Bucket",
  "planId": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
  "orderHint": "8585363889524958496",
  "id": "ttEB_Uj690STdR3GC1MIDZgANq1U"
}
```
