# spo term list

Lists taxonomy terms from the given term set

## Usage

```sh
m365 spo term list [options]
```

## Options

`--termGroupId [termGroupId]`
: ID of the term group where the term set is located. Specify `termGroupId` or `termGroupName` but not both

`--termGroupName [termGroupName]`
: Name of the term group where the term set is located. Specify `termGroupId` or `termGroupName` but not both

`--termSetId [termSetId]`
: ID of the term set for which to retrieve terms. Specify `termSetId` or `termSetName` but not both

`--termSetName [termSetName]`
: Name of the term set for which to retrieve terms. Specify `termSetId` or `termSetName` but not both

`--includeChildTerms`
: If specified, child terms are loaded as well.

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

List taxonomy terms from the term group and term set with the given name

```sh
m365 spo term list --termGroupName PnPTermSets --termSetName PnP-Organizations
```

List taxonomy terms from the term group and term set with the given ID

```sh
m365 spo term list --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termSetId 0e8f395e-ff58-4d45-9ff7-e331ab728bec
```

List taxonomy terms from the term group and term set with the given ID including child terms if any are found

```sh
m365 spo term list --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termSetId 0e8f395e-ff58-4d45-9ff7-e331ab728bec --includeChildTerms
```

## Response

### Standard response

=== "JSON"

    ```json
    [
      {
        "_ObjectType_": "SP.Taxonomy.Term",
        "_ObjectIdentity_": "430486a0-200a-6000-02cc-2eb89d8dd424|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSNsaKRf7N4K0qppo4zwLIRZg==",
        "CreatedDate": "2021-10-12T09:42:44.880Z",
        "Id": "7f91a2b1-78b3-4a2b-a9a6-8e33c0b21166",
        "LastModifiedDate": "2021-10-12T09:42:45.237Z",
        "Name": "Departments",
        "CustomProperties": {},
        "CustomSortOrder": null,
        "IsAvailableForTagging": true,
        "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
        "Description": "",
        "IsDeprecated": false,
        "IsKeyword": false,
        "IsPinned": false,
        "IsPinnedRoot": false,
        "IsReused": false,
        "IsRoot": true,
        "IsSourceTerm": true,
        "LocalCustomProperties": {
          "Id": "termDepartments"
        },
        "MergedTermIds": [],
        "PathOfTerm": "Departments",
        "TermsCount": 1
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Name      
    ------------------------------------  ---------
    c387e91c-b553-4b92-886b-9af717cd73b0  Financing
    ```

=== "CSV"

    ```csv
    Id,Name
    c387e91c-b553-4b92-886b-9af717cd73b0,Financing
    ```

### `includeChildTerms` response

When we make use of the option `includeChildTerms` the response will differ. 

=== "JSON"

    ```json
    [
      {
        "_ObjectType_": "SP.Taxonomy.Term",
        "_ObjectIdentity_": "430486a0-200a-6000-02cc-2eb89d8dd424|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSNsaKRf7N4K0qppo4zwLIRZg==",
        "CreatedDate": "2021-10-12T09:42:44.880Z",
        "Id": "7f91a2b1-78b3-4a2b-a9a6-8e33c0b21166",
        "LastModifiedDate": "2021-10-12T09:42:45.237Z",
        "Name": "Departments",
        "CustomProperties": {},
        "CustomSortOrder": null,
        "IsAvailableForTagging": true,
        "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
        "Description": "",
        "IsDeprecated": false,
        "IsKeyword": false,
        "IsPinned": false,
        "IsPinnedRoot": false,
        "IsReused": false,
        "IsRoot": true,
        "IsSourceTerm": true,
        "LocalCustomProperties": {
          "Id": "termDepartments"
        },
        "MergedTermIds": [],
        "PathOfTerm": "Departments",
        "TermsCount": 1,
        "Children": [{
          "_ObjectType_": "SP.Taxonomy.Term",
          "_ObjectIdentity_": "d10486a0-4067-5000-de97-bf1ab5bff53c|fec14c62-7c3b-481b-851b-c80d7802b224:te:kTm3XibpGUiE5nxBtVMTf14Jch8b6X1EtvEo9yq4/mCesjVWlBPHRaBqFOZeTRSNf5ybB0mfG0Kx6YO5CODZ1A==",
          "CreatedDate": "2022-12-25T23:59:15.200Z",
          "Id": "079b9c7f-9f49-421b-b1e9-83b908e0d9d4",
          "LastModifiedDate": "2022-12-25T23:59:15.200Z",
          "Name": "Financing",
          "CustomProperties": {},
          "CustomSortOrder": null,
          "IsAvailableForTagging": true,
          "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
          "Description": "",
          "IsDeprecated": false,
          "IsKeyword": false,
          "IsPinned": false,
          "IsPinnedRoot": false,
          "IsReused": false,
          "IsRoot": false,
          "IsSourceTerm": true,
          "LocalCustomProperties": {},
          "MergedTermIds": [],
          "PathOfTerm": "Departments;Financing",
          "TermsCount": 0
        }]
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Name       ParentTermId
    ------------------------------------  ---------  ------------------------------------
    c387e91c-b553-4b92-886b-9af717cd73b0  Financing  079b9c7f-9f49-421b-b1e9-83b908e0d9d4
    ```

=== "CSV"

    ```csv
    Id,Name,ParentTermId
    c387e91c-b553-4b92-886b-9af717cd73b0,Financing,079b9c7f-9f49-421b-b1e9-83b908e0d9d4
    ```

