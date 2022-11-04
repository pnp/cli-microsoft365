# pp gateway list

Returns a list of gateways for which the user is an admin

## Usage

```sh
m365 pp gateway list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

List gateways for which the user is an admin

```sh
m365 pp gateway list
```

## Response

=== "JSON"

    ```json
    [
      {
        "id":"22660b34-31b3-4744-a99c-5e154458a784",
        "gatewayId":0,
        "name":"Contoso Gateway",
        "type":"Resource",
        "publicKey":{
          "exponent":"AQAB",
          "modulus":"okJBN8MJyaVkjfkN75B6OgP7RYiC3KFMFaky9KqqudqiTOcZPRXlsG+emrbnnBpFzw7ywe4gWtUGnPCqy01RKeDZrFA3QfkVPJpH28OWfrmgkMQNsI4Op2uxwEyjnJAyfYxIsHlpevOZoDKpWJgV+sH6MRf/+LK4hN3vNJuWKKpf90rNwjipnYMumHyKVkd4Vssc9Ftsu4Samu0/TkXzUkyje5DxMF2ZK1Nt2TgItBcpKi4wLCP4bPDYYaa9vfOmBlji7U+gwuE5bjnmjazFljQ5sOP0VdA0fRoId3+nI7n1rSgRq265jNHX84HZbm2D/Pk8C0dElTmYEswGPDWEJQ=="
        },
        "gatewayAnnotation":"{\"gatewayContactInformation\":[\"admin@contoso.onmicrosoft.com\"],\"gatewayVersion\":\"3000.122.8\",\"gatewayWitnessString\":\"{\\\"EncryptedResult\\\":\\\"UyfEqNSy0e9S4D0m9oacPyYhgiXLWusCiKepoLudnTEe68iw9qEaV6qNqTbSKlVUwUkD9KjbnbV0O3vU97Q/KTJXpw9/1SiyhpO+JN1rcaL51mPjyQo0WwMHMo2PU3rdEyxsLjkJxJZHTh4+XGB/lQ==\\\",\\\"IV\\\":\\\"QxCYjHEl8Ab9i78ZBYpnDw==\\\",\\\"Signature\\\":\\\"upVXK3DvWdj5scw8iUDDilzQz1ovuNgeuXRpmf0N828=\\\"}\",\"gatewayMachine\":\"SPFxDevelop\",\"gatewaySalt\":\"rA1M34AdgdCbOYQMvo/izA==\",\"gatewayWitnessStringLegacy\":null,\"gatewaySaltLegacy\":null,\"gatewayDepartment\":null,\"gatewayVirtualNetworkSubnetId\":null}"
      }
    ]
    ```

=== "Text"

    ```text
    id                                   name
    ------------------------------------ ---------------
    22660b34-31b3-4744-a99c-5e154458a784 Contoso Gateway
    ```

=== "CSV"

    ```csv
    id,name
    22660b34-31b3-4744-a99c-5e154458a784,Contoso Gateway
    ```
