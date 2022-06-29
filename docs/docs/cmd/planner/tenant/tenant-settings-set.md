# planner tenant settings set

Sets Microsoft Planner configuration of the tenant

## Usage

```sh
m365 planner tenant settings set [options]
```

## Options

`--isPlannerAllowed [isPlannerAllowed]`
: Configure whether Planner should be enabled on the tenant.

`--allowCalendarSharing [allowCalendarSharing]`
: Configure whether Outlook calendar sync is enabled.

`--allowTenantMoveWithDataLoss [allowTenantMoveWithDataLoss]`
: Configure whether a tenant move into a new region is authorized.

`--allowTenantMoveWithDataMigration [allowTenantMoveWithDataMigration]`
: Configure whether a tenant move with data migration is authorized.

`--allowRosterCreation [allowRosterCreation]`
: Configure whether Planner roster creation is allowed.

`--allowPlannerMobilePushNotifications [allowPlannerMobilePushNotifications]`
: Configure whether push notifications are enabled in the mobile app.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a global administrator.

## Examples

Disable Microsoft Planner in the tenant

```sh
m365 planner tenant settings set --isPlannerAllowed false
```

Disable Outlook calendar sync and mobile push notifications

```sh
m365 planner tenant settings set --allowCalendarSharing false --allowPlannerMobilePushNotifications false
```

Enable Microsoft Planner but disallow roster plans to be created

```sh
m365 planner tenant settings set --isPlannerAllowed true --allowRosterCreation false
```
