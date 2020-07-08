# Communication with Microsoft 365

To manage settings of the different Microsoft 365 services, the CLI for Microsoft 365 uses REST APIs exposed by the corresponding services. Using the REST APIs is meant to promote consistency and reusability of code and tests across the CLI no matter which Microsoft 365 service the CLI is communicating with.

Some SharePoint Online commands deviate from this rule and mimic SharePoint CSOM calls instead. This is done out of necessity as some operations, such as managing Microsoft 365 CDN settings or tenant properties, are not exposed through REST APIs. Whenever REST APIs become available for these operations, the affected commands will be changed to use REST APIs instead of mimicking CSOM calls.
