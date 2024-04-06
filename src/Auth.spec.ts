import { DeviceCodeResponse } from "@azure/msal-common";
import * as msal from '@azure/msal-node';
import assert from 'assert';
import clipboard from 'clipboardy';
import fs from 'fs';
import 'node-forge';
import sinon from 'sinon';
import { Auth, AuthType, CertificateType, CloudType, Connection, InteractiveAuthorizationCodeResponse, InteractiveAuthorizationErrorResponse } from './Auth.js';
import authServer from './AuthServer.js';
import { CommandError } from './Command.js';
import { FileTokenStorage } from './auth/FileTokenStorage.js';
import { TokenStorage } from './auth/TokenStorage.js';
import { cli } from './cli/cli.js';
import { Logger } from './cli/Logger.js';
import request from './request.js';
import { accessToken } from "./utils/accessToken.js";
import { browserUtil } from "./utils/browserUtil.js";
import { sinonUtil } from './utils/sinonUtil.js';

class MockTokenStorage implements TokenStorage {
  public get(): Promise<string> {
    return Promise.resolve('ABC');
  }

  public set(): Promise<void> {
    return Promise.resolve();
  }

  public remove(): Promise<void> {
    return Promise.resolve();
  }
}

const mockTokenCachePlugin: msal.ICachePlugin = {
  beforeCacheAccess(tokenCacheContext: msal.TokenCacheContext): Promise<void> {
    tokenCacheContext.tokenCache.deserialize('');
    return Promise.resolve();
  },
  afterCacheAccess(): Promise<void> {
    return Promise.resolve();
  }
};

describe('Auth', () => {
  let log: any[];
  let auth: Auth;
  let response: DeviceCodeResponse;
  let openStub: sinon.SinonStub;
  let clipboardStub: sinon.SinonStub;
  let getSettingWithDefaultValueStub: sinon.SinonStub;
  const resource: string = 'https://contoso.sharepoint.com';
  let loggerSpy: sinon.SinonSpy;
  const logger: Logger = {
    log: (msg: any) => log.push(msg) as any,
    logRaw: (msg: any) => log.push(msg) as any,
    logToStderr: (msg: any) => log.push(msg) as any
  };
  const loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  let readFileSyncStub: sinon.SinonStub;
  let initializeServerStub: sinon.SinonStub;
  const publicApplication = new msal.PublicClientApplication({
    auth: {
      clientId: '123'
    },
    cache: {
      cachePlugin: mockTokenCachePlugin
    }
  });
  const identityName = 'someuser';
  const identityId = '34b70d68-17b0-4b54-b2dd-8f85ebc9d624';
  const identityTenantId = '9bc3ab49-b65d-410a-85ad-de819febfddd';
  const appId = '9bc3ab49-b65d-410a-85ad-de819febfddc';
  const tenant = '9bc3ab49-b65d-410a-85ad-de819febfddd';
  const activeConnection: Connection = { name: identityId, identityId, identityName, active: true, appId, tenant, authType: AuthType.DeviceCode, certificateType: CertificateType.Unknown, accessTokens: {}, cloudType: CloudType.Public, identityTenantId: identityTenantId, deactivate: () => { } };
  const base64EncodedPemCert = 'QmFnIEF0dHJpYnV0ZXMNCiAgICBsb2NhbEtleUlEOiBDQyBGNCBGMiBBMyBDMyBEMiAwOSBDNSAxMiBCMyA3MiA0QiBCOCA4MyBBNSA0NyA0QyAwOSAyMSBEQyANCnN1YmplY3Q9QyA9IEFVLCBTVCA9IFNvbWUtU3RhdGUsIE8gPSBJbnRlcm5ldCBXaWRnaXRzIFB0eSBMdGQNCg0KaXNzdWVyPUMgPSBBVSwgU1QgPSBTb21lLVN0YXRlLCBPID0gSW50ZXJuZXQgV2lkZ2l0cyBQdHkgTHRkDQoNCi0tLS0tQkVHSU4gQ0VSVElGSUNBVEUtLS0tLQ0KTUlJRGF6Q0NBbE9nQXdJQkFnSVVXb25VNFM0RTcxRjVZMU5zU0xYbUlhZ1dkNVl3RFFZSktvWklodmNOQVFFTA0KQlFBd1JURUxNQWtHQTFVRUJoTUNRVlV4RXpBUkJnTlZCQWdNQ2xOdmJXVXRVM1JoZEdVeElUQWZCZ05WQkFvTQ0KR0VsdWRHVnlibVYwSUZkcFpHZHBkSE1nVUhSNUlFeDBaREFlRncweE9UQTNNVEl5TVRVek1qbGFGdzB5TURBMw0KTVRFeU1UVXpNamxhTUVVeEN6QUpCZ05WQkFZVEFrRlZNUk13RVFZRFZRUUlEQXBUYjIxbExWTjBZWFJsTVNFdw0KSHdZRFZRUUtEQmhKYm5SbGNtNWxkQ0JYYVdSbmFYUnpJRkIwZVNCTWRHUXdnZ0VpTUEwR0NTcUdTSWIzRFFFQg0KQVFVQUE0SUJEd0F3Z2dFS0FvSUJBUUNsa01lQXlKbTJkMy95aEV0NHZGYjYrYjEyUGxRSDB4VGx1a1BoK2xScg0KOXJDNk5DM3dObnoySm5vbE1HclhuZVp2TlN5czFONVpSTm0yTjhQdy9QOExxeHJSenFFOFBNVC96NnN1UFhSUg0KWm5hZ2xaUklXb0NNR25pRVlDZVJHZnI4R2JpUXcwYlZEeXFuSnJaZjByS0pHbnZUNlY3QmpUdFloRWIzeXhoNA0KSmNUSnIrVDl0OEFYaldmemt6alBZdklxYmhha3FxcHd1SEVPYkh4T201cHVERTFBNVJOZm8wamcrTmZtVko5VQ0KMWR1RjVzdmE2NVQ5Q1RtdEdlbVNlUGlzWmgxZmhoOS94QmJwTCs0RUJWUXZqdEZXWk5zMVJHMW9QUllscmpzaQ0KTXFsaHNUdjhDZXI5cWUxcVNTdHFjMmJsc3hGek1zNmxZOHAvUHIrYm5uR3pBZ01CQUFHalV6QlJNQjBHQTFVZA0KRGdRV0JCU203cWFreXQwY2xxN0lnRFRWdkUrWEpaNFU5akFmQmdOVkhTTUVHREFXZ0JTbTdxYWt5dDBjbHE3SQ0KZ0RUVnZFK1hKWjRVOWpBUEJnTlZIUk1CQWY4RUJUQURBUUgvTUEwR0NTcUdTSWIzRFFFQkN3VUFBNElCQVFBYQ0KQnVqTytveU0yL0Q0SzNpS3lqVDVzbHF2UFVlVzFrZVVXYVdSVDZXRTY0VkFPbTlPZzU1bkIyOE5TSVVXampXMA0KdTJEUHF3SzJiOEFXalEveWp3S3NUMXVTdzcyQ0VEY2o3SkE1VXA5UWpBa0hIZmFoQWtOd0o5M0llcmFBdTEyVQ0KN25FRDdIN20yeGZscDVwM0dadzNHUE0rZmpBaDZLOUZIRDI0bWdGUTh4b2JPQSttVEVvV2ZIVVQrZ1pUMGxYdQ0KazFrVTJVelVOd2dwc3c4V04wNFFzWU5XcFF5d3ppUWtuZTQzNW5tdmxZOGZRc2hPSnErK0JCS0thd0xEcjk3bA0KRTBYQUxEZDZlVVhQenZ5OU1xZlozeUswRmUzMy8zbnZnUnE4QWZ3azRsbzhac2ZYWUlSTXA3b3BER0VmaUZmNQ0KM3JTTGxSZG9TNDQ4OVFZRnAyYUQNCi0tLS0tRU5EIENFUlRJRklDQVRFLS0tLS0NCkJhZyBBdHRyaWJ1dGVzDQogICAgbG9jYWxLZXlJRDogQ0MgRjQgRjIgQTMgQzMgRDIgMDkgQzUgMTIgQjMgNzIgNEIgQjggODMgQTUgNDcgNEMgMDkgMjEgREMgDQpLZXkgQXR0cmlidXRlczogPE5vIEF0dHJpYnV0ZXM+DQotLS0tLUJFR0lOIFBSSVZBVEUgS0VZLS0tLS0NCk1JSUV2Z0lCQURBTkJna3Foa2lHOXcwQkFRRUZBQVNDQktnd2dnU2tBZ0VBQW9JQkFRQ2xrTWVBeUptMmQzL3kNCmhFdDR2RmI2K2IxMlBsUUgweFRsdWtQaCtsUnI5ckM2TkMzd05uejJKbm9sTUdyWG5lWnZOU3lzMU41WlJObTINCk44UHcvUDhMcXhyUnpxRThQTVQvejZzdVBYUlJabmFnbFpSSVdvQ01HbmlFWUNlUkdmcjhHYmlRdzBiVkR5cW4NCkpyWmYwcktKR252VDZWN0JqVHRZaEViM3l4aDRKY1RKcitUOXQ4QVhqV2Z6a3pqUFl2SXFiaGFrcXFwd3VIRU8NCmJIeE9tNXB1REUxQTVSTmZvMGpnK05mbVZKOVUxZHVGNXN2YTY1VDlDVG10R2VtU2VQaXNaaDFmaGg5L3hCYnANCkwrNEVCVlF2anRGV1pOczFSRzFvUFJZbHJqc2lNcWxoc1R2OENlcjlxZTFxU1N0cWMyYmxzeEZ6TXM2bFk4cC8NClByK2Jubkd6QWdNQkFBRUNnZ0VBUjRsMytqZ3kybmxseWtiSlNXQ3ZnSCs2RWtZNkRxdHd3eFlwVUpIV09sUDcNCjVtaTNWS3htY0FFT0U5V0l4S05RTnNyV0E5TnlRMFlSZjc4MnBZRGJQcEp1NHlxUjFqSTN1SVJsWlhSZU52RzcNCjNnVGpiaVBVbVRTeTBCZXY0TzFGMmZuUEdwV1ZuR2VTT1dqcnNobWExTXlocGwyV2VMRHFiSU96R2t3aHhYOXkNClRhRFd5MjErbDFpNVNGWUZTdHdXOWlhOXRORTFTTTU4WnpQWk0yK0NDdHhQVEFBQXRJRmZXUVdTbnhodUxMenMNCjNyVDRVOGNLZzJITVBXb29rOS9peWxsa0xEVXBPanhJR2tHWXdheDVnR2xvR0xZYWVoelc5Q3hobzgvc3A4WjUNCkVNNVFvczVJSTF2K21pNHhHa0RTdW4rbDYzcDN5Nm54T3pqM1h1MzRlUUtCZ1FEUDNtRWttN2lVaTlhRUxweXYNCkIxeDFlRFR2UmEwcllZMHZUaXFrYzhyUGc0NU1uOUNWRWZqdnV3YkN4M21tTExabThqZVY3ZTFHWjZJeXYreEUNCmcxeFkrUTd0RUlCb1FwWThlemg0UVYvMXRkZkhiUzNPcGdIbHVqMGd5MWxqT2QrbkxzS2RNQWRlYVF3Uy9WK2MNCk51Sks0Y3oyQWl6UXU1dHQ4WHdoOGdvU0Z3S0JnUURMNXRjZnF0VmdMQWJmMnJQbEhBLzdNcU1sWGpqNUQ0ejkNCjZmTWlCVDdOWHlYUGx6a2pJQkxOdG9OWlBCVTFzeERFb2tiNUtyTlhLTUtIaU9nTkQ0cWtDYkdnRFk2WUdaS3cNCkg4bDlLWDBaM2pwcEp0TURvQ21yQW9hSmZTUXNreGJXSDd4VlFGVzdPVWQ0dHMxZ3FDbTBUTFVxeW9lcW1EK3INCmg3WFlaa2RxeFFLQmdBK2NpZnN2M3NyNVBhRXJ4d1MyTHRGN3Q2NElzNXJBZHRRSXNOY3RBeHhXcXdkQ01XNGcNCnJXdUR4bHcya3dKUjlWa0I4LzdFb2I5WjVTcWVrMllKMzVPbkVPSHBEVnZITkhWU1k4bFVUNXFxajR3Z3ZRSDYNCkljWlpHR0l3STRSNlFqdlNIVGVrOWNpM1p2cStJTUlndFJvZW4wQVNwYjcvZUFybnlnVGFvcnI5QW9HQkFJT3QNCllOSEhqaUtjYkJnV2NjU01tZGw4T3hXL3dvVTlRSzBkYjNGUjk5dkREWFVCVU5uWk5hdDVxVnR3VExZd0hLMFANCnEwdndBbjlRQ0VoazVvN0FzYVQ3eWFUMS9GZEhkSTZmQ0l6MnhSNTJnRHcxNFdIZkJlbTFLTk1UYU5BTWNWdjQNCmhMUjlacUFRL3BIN1k2aC9FT2VwL2ZsVGI4ZUFxT1dLTDZvL2F2R05Bb0dCQUlHc0c1VExuSmlPU044SUtGU04NCmJmK3IrNkhWL2R6MkluNjhSR255MTB0OGpwbUpPbGgrdXRncGtvOXI2Y09uWGY4VHM2SFAveTBtbDl5YXhvMlANCm52c2wwcFlseFQxQy9taXJaZWxYKzFaQTltdFpHT2RxbzZhdVZUM1drcXBpb3c2WUtzbzl2Z2RHWmRWRUxiMEINCnUvdyt4UjBvN21aSEpwVEdmS09KdE53MQ0KLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLQ0K';
  const base64EncodedPfxCert = 'MIIJqQIBAzCCCW8GCSqGSIb3DQEHAaCCCWAEgglcMIIJWDCCBA8GCSqGSIb3DQEHBqCCBAAwggP8AgEAMIID9QYJKoZIhvcNAQcBMBwGCiqGSIb3DQEMAQYwDgQIzLm7KYappOYCAggAgIIDyPpygKYYXv/M6WX6QGX/ltZYjTCM/OSpzmHrBwho+e1ZgPXKsxi+P4tU31g+B0HFT2tVtpKULzu3NHxs2nzfWW9POomI8NSK4AC+yPnC7qVkcL+6pwW9kDACXS6xyY3i6kRevBPz1BZ09BPiR4VQBl+5r1AhraIc1mEMOnUljNO1tj7sN9tyQYuzNGXGsJ/WdVzIGg27LM2BkiP0Mo5933Pk5sg/Y1+fEiPNNa0VdoPWmpFGZ1t16p13tUGzzcwaj4oxYTpu7C25GY9xZ/HidlPqRsUWj29VtFo+Yzo+uYQRkV7VcT3oBa0If60Yw3G5xYrW+Qf+Y2CMG6nKLYLsh5J0yGSTEOG4s6JiKk7O1YQHghzAEiPi9Oe/inyFUjc+DYXcIWnIS/uw2GjgTBETnvV5ftMJrmkBvfSiT72pBGjXji41dPscAA7NohsVNCzQYGJvWWG8B/BnWp6VJuh91Aerq8fSg6K/oc44CAvFdYrOHm87xWG4nPlURIIuqBCm1DDMYLB8rgRhWAcOxpTDruj0X5Ve/X5sNCORlD6M2sxFC8ictLI3pv6ZYlDFxvIBOHUBhXxXg5x8xmNixALmQSBrQUj7uMD71qjtyMSNW/ow+S/fZqxzU8z6CSncYDHaWH1+HJhjxpC62u2cyYQXqBCJZ44cT6gZKRIt4HxEph8hiQMAcXjLyu91IGZjCPB3FbPgqFjzc3LUojj38DSQxF9Oo6BKOcMls4fZc8sdipF7pJLBgxXmrdwyy6Ge7VtewblgOuW2n+7MneNDsbIyfssNiO2aDp+SfBNT5fEhzv3gH3AdW25RByiG1EJJBP+ZQolM6AfWxJFRibCySlZPkgYT9RgqCtI4hH068KEan1sX8VLl/M838bOdiFHPyDMw7/5HZu6jFVjiMTXO3ry7M0kDaHLNgt0cDQqEwAZ/pWEamlwR3/vY+Ofgy1cFchaxz4MPQYer214+77N65GcIxn7D3biqLCVVhglUdJvFBH8JqaKrmlGYxL8sFuBp5mBGdGQcEdRvEr1sSMWE2hdYRfkBfVIn3eTPkTSL2J6d1FV8DKH0tNuWqY+W/fjwK2w+WF8iiCgtKMVQYPp/RoXZCxHaweEqi2icrB3J9HWzHpSpIdvghrgwAe87UpbwYdBonsW0EbYv9GeDaWasI8JTYt6WHN7cQVIlVdI0hrqJ4e5aEUWyU22CjDp4M9RrvVge7UDFAAF3KbEc3e6H39frb6GnovjIpW/40eAIUpuOTtgDSxUpI8tulp7pTDXvaH8oElrns5e9leoHMIIFQQYJKoZIhvcNAQcBoIIFMgSCBS4wggUqMIIFJgYLKoZIhvcNAQwKAQKgggTuMIIE6jAcBgoqhkiG9w0BDAEDMA4ECPEeujz28p7JAgIIAASCBMjGEjCHGk8FZXleYoXwd/P3Hml08yliW3jZ+50ynrheZDe7F2d2QdValQuS/YGF1B1pnSsIT/E9cu3n2S2QqCVPNNjd2I58SmB+uoOAj9Ng57y1RFQr4BFMxhEmjnKcxtbr95v8B2hxesKvXmVj3QhvNNHApaYEZ6LlL2xJxQpN1aCEIWPoOOq1uJrDkPwjB7vyt1OE6+v1wTy6DN9gurBR6KYnFgf+/6HQDW3YcfNLBwGC9/KBXvGmzBm/LBKNeDUYReXDpgNxnWhWX6t3sHhrkGNhp4r/Ds3uN+sN8JhQXZ6Fncu8OHBuou9KQKwQSpWsxqIb7IQF/B07FI0d1ahq12GlqnUrzB0nzsDKFioxvLsV3IBuKRxAEMDngo+6HnnTpVLK2qhLjaB8+38lpQv8mfVbugGIOcyBSVUGYDwXoBU9Q/8RXYO1D9l90MU9j9VWz22HidtrosFR9iIfYCupwx/WiTvJMbUHj8glpq7nd3cIWhCbxlb57AsXx9r+GnEOGmiaESNO1NCN5HpluWRzdjOUVQY6K54QG9n8M3GgKoAibWA66bL/UgAx/neiyqcGFWlTdQpuY/ZdDKq6CmBpm+emu6Fj9j8awvbc53tvJCnvEAluo/eB4nOTcNXFzVKpPzMT8GwNY9YoU3m9WX3sPWdgk3U/+ij1EyW93bjhINFxwlvHtIPDdKt1g3pM/QYZnG3/bOUmZRNltlxRvNTFdqBwuQQYcTTyHSgDvKnpTCEPLH+fnaQ5oIDSf2olYT4O9ALKvC+3y5eodrBZIciZX9TSP65BRfQShW0XIDgtGv5bu8DZwiRUVf6QvRbyySkx8NdqxNG4s5U+PiF++jj/X89EuwNjZqtjuejoNqGfWpxhwIdUaAdhvnrq+KToA3V+WotZHrYwkkrmvpYr48dteCrdDw92drQyrgsanMev5qngXUZLHJFFxf+kJ2DhMF+XjLOWTLYK/daJ0FATWAMrclY7petJTDEDOx1qJu+l3BEZ6yKwQ5v/bicDDvx7JBi3KbIHk4zuW9LXhxdhRCAZMPXARjBo6IEie7+Jw7N8HPVa6VtTKZiFVbfzHvsie0sD648qBNHqm5mPzXnNlf8ok5WPXvW9vdHKo6nHl7NANUkXEwSjXV/v15ATfyHQQivxLIlWrBSiepRS1LvtWwybTpvD781DaesvLSqJLLP1tGoLUBYE1vQ3/zTe2psBVFbmw3IHCrVEPAaduVTUeB2UIxYWwJlwe4hIlu+cPHCrUlayOS4qB0RliHX9xAmGrpjxuvAk+M5r7m2+KLq4Rkv6ITrlpRkhO8dCD5hmE0y5qRVGpv107fL0K+ya8l3sJVIacfG/qYoaTzqn896gXnR/aURD+XdaAl1JCAV2K64H8wU3cNwwbFoDB+qhBpXogHmW+XgTBuSJoR2/6vZ7G9w6Ht949WeUpzsmtRsSj+c+kz1rBnRDHT9nykB3xwtghINhwcHumhMkTK87EKJ+mAM9hRLVGTsOlxir+0DhS7JwhKSHOVcAjnMf3Nf5jpPGrWxZQD9ppqMut4M5GE8mbSRR8bPa/H9//0Y0hW5ALwaCIWVht+h3rk0m8wb7gJZYkMktOgbWX5kmYEzuJb3zptGIKY/siD3fJLcxJTAjBgkqhkiG9w0BCRUxFgQUzPTyo8PSCcUSs3JLuIOlR0wJIdwwMTAhMAkGBSsOAwIaBQAEFKgCEPptVqSh/raIMgRw+Ixd0qrTBAiptv/LHThdywICCAA=';
  const tokenCache = {
    getAllAccounts: () => Promise.resolve([{ localAccountId: identityId }]),
    getAccountByLocalId: (localAccountId: string) => Promise.resolve({ localAccountId: localAccountId }),
    removeAccount: () => Promise.resolve()
  };

  const httpServerResponse = <InteractiveAuthorizationCodeResponse>{
    code: "secretCode",
    redirectUri: "https://localhost:666"
  };

  before(() => {
    sinon.stub(publicApplication, 'getTokenCache').callsFake(() => tokenCache as any);
  });

  beforeEach(() => {
    log = [];
    auth = new Auth();
    response = {
      deviceCode: "",
      expiresIn: 0,
      interval: 0,
      message: "",
      userCode: "",
      verificationUri: ""
    };
    auth.connection.active = true;
    auth.connection.identityId = identityId;
    auth.connection.identityName = identityName;
    auth.connection.appId = appId;
    auth.connection.tenant = tenant;
    (auth as any)._authServer = authServer;
    readFileSyncStub = sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
    initializeServerStub = sinon.stub((auth as any)._authServer, 'initializeServer').callsFake(((connection: Connection, resource: string, resolve: (error: InteractiveAuthorizationCodeResponse) => void) => {
      resolve(httpServerResponse);
    }) as any);
    loggerSpy = sinon.spy(logger, 'log');
    (auth as any)._clipboardy = clipboard;
    openStub = sinon.stub(browserUtil, 'open').callsFake(async () => { return; });
    clipboardStub = sinon.stub((auth as any)._clipboardy, 'writeSync').callsFake(() => 'clippy');
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => 'key'));
    sinon.stub(auth as any, 'getConnectionInfoFromStorage').resolves(activeConnection);
    sinon.stub(auth as any, 'getAllConnectionsFromStorage').resolves([activeConnection]);
    sinon.stub(auth, 'storeConnectionInfo').resolves();
  });

  afterEach(() => {
    loggerSpy.restore();
    readFileSyncStub.restore();
    initializeServerStub.restore();
    sinonUtil.restore([
      cli.getConfig().get,
      request.get,
      (auth as any).getClientApplication,
      (auth as any).getDeviceCodeResponse,
      (auth as any).storeConnectionInfo,
      (auth as any).getConnectionInfoFromStorage,
      (auth as any).getAllConnectionsFromStorage,
      publicApplication.acquireTokenSilent,
      publicApplication.acquireTokenByDeviceCode,
      publicApplication.acquireTokenByUsernamePassword,
      publicApplication.acquireTokenByCode,
      tokenCache.getAllAccounts
    ]);
    openStub.restore();
    clipboardStub.restore();
    getSettingWithDefaultValueStub.restore();
  });

  after(() => {
    auth.connection.deactivate();
  });

  it('returns existing access token if still valid', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() + 1);
    auth.connection.accessTokens[resource] = {
      expiresOn: now.toISOString(),
      accessToken: 'abc'
    };

    auth.ensureAccessToken(resource, logger).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, auth.connection.accessTokens[resource].accessToken);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('returns existing access token if still valid (debug)', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() + 1);
    auth.connection.accessTokens[resource] = {
      expiresOn: now.toISOString(),
      accessToken: 'abc'
    };

    auth.ensureAccessToken(resource, logger, true).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, auth.connection.accessTokens[resource].accessToken);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('returns existing access token if still valid (token stored as a Date)', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() + 1);
    auth.connection.accessTokens[resource] = {
      expiresOn: now,
      accessToken: 'abc'
    };

    auth.ensureAccessToken(resource, logger).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, auth.connection.accessTokens[resource].accessToken);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves new access token silently if already signed in', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    const acquireTokenSilentStub = sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'abc'
    } as any));

    auth.ensureAccessToken(resource, logger).then(() => {
      try {
        assert(acquireTokenSilentStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves new access token silently if already signed in (debug)', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    const acquireTokenSilentStub = sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'abc'
    } as any));

    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert(acquireTokenSilentStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when retrieving new access token silently', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.reject('An error has occurred'));

    auth.ensureAccessToken(resource, logger).then(() => {
      done('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving new access token silently (debug)', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.reject({
      errorCode: 'error',
      errorMessage: 'AADSTS00000 An error has occurred'
    }));

    auth.ensureAccessToken(resource, logger).then(() => {
      done('Got access token');
    }, (err) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure response when retrieving new access token', (done) => {
    auth.connection.certificate = base64EncodedPemCert;
    auth.connection.authType = AuthType.Certificate;

    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve(null));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(() => {
      done('Got access token');
    }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('getTokenPromise authentication result is null.'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles empty response when retrieving new access token', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.reject('An error has occurred'));

    auth.ensureAccessToken(resource, logger, true).then(() => {
      done('Got access token');
    }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('getTokenPromise authentication result is null.'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows AAD error when retrieving new access token silently failed', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.reject({
      errorCode: 'error',
      errorMessage: 'AADSTS00000 An error has occurred'
    }));

    auth.ensureAccessToken(resource, logger).then(() => {
      done('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify({
          errorCode: 'error',
          errorMessage: 'AADSTS00000 An error has occurred'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when invalid Microsoft Entra app used', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    sinon.stub(publicApplication, 'acquireTokenByDeviceCode').callsFake(_ => Promise.reject({
      errorCode: 'invalid_client',
      errorMessage: `AADSTS7000218: The request body must contain the following parameter: 'client_assertion' or 'client_secret'.`
    }));

    auth.ensureAccessToken(resource, logger).then(() => {
      done('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify({
          errorCode: 'invalid_client',
          errorMessage: `AADSTS7000218: The request body must contain the following parameter: 'client_assertion' or 'client_secret'.`
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves new access token using existing refresh token when the access token expired (debug)', (done) => {
    const now = new Date();
    now.setSeconds(now.getSeconds() - 1);
    auth.connection.accessTokens[resource] = {
      expiresOn: now.toISOString(),
      accessToken: 'abc'
    };
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));

    auth.ensureAccessToken(resource, logger, true).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, 'acc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves new access token using existing refresh token when refresh forced', (done) => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').callsFake((() => { }) as any);
    const now = new Date();
    now.setSeconds(now.getSeconds() + 1);
    auth.connection.accessTokens[resource] = {
      expiresOn: now.toISOString(),
      accessToken: 'abc'
    };
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(publicApplication, 'acquireTokenSilent').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));

    auth.ensureAccessToken(resource, logger, true, true).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, 'acc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves access token using device code authentication flow when no refresh token available and no authType specified', (done) => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').callsFake((() => 'value'));
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    sinon.stub(publicApplication, 'acquireTokenByDeviceCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));

    auth.ensureAccessToken(resource, logger).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, 'acc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('opens the browser with the login (using autoOpenLinksInBrowser)', async () => {
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => {
      if (settingName === "autoOpenLinksInBrowser") { return true; }
      else {
        return defaultValue;
      }
    }));

    openStub.restore();
    openStub = sinon.stub(browserUtil, 'open').callsFake(async () => { return; });

    await (auth as any).processDeviceCodeCallback(response, logger, false);
    assert(openStub.called);
  });

  it('copies the device code to the clipboard', async () => {
    await (auth as any).processDeviceCodeCallback(response, logger, false);
    assert(clipboardStub.called);
  });

  it('writes response from the device code request (debug)', async () => {
    await (auth as any).processDeviceCodeCallback(response, logger, true);
    assert(loggerLogToStderrSpy.calledWith(response));
  });

  it('retrieves token using device code authentication flow when authType deviceCode specified', (done) => {
    const config = cli.getConfig();
    sinon.stub(config, 'get').callsFake((() => 'value'));
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    const acquireTokenByDeviceCodeStub = sinon.stub(publicApplication, 'acquireTokenByDeviceCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));
    auth.connection.authType = AuthType.DeviceCode;

    auth.ensureAccessToken(resource, logger).then(() => {
      try {
        assert(acquireTokenByDeviceCodeStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using password flow when authType password specified', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    const acquireTokenByUsernamePasswordStub = sinon.stub(publicApplication, 'acquireTokenByUsernamePassword').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));
    auth.connection.authType = AuthType.Password;

    auth.ensureAccessToken(resource, logger).then(() => {
      try {
        assert(acquireTokenByUsernamePasswordStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using password flow when authType password specified (debug)', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    const acquireTokenByUsernamePasswordStub = sinon.stub(publicApplication, 'acquireTokenByUsernamePassword').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));
    auth.connection.authType = AuthType.Password;

    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert(acquireTokenByUsernamePasswordStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when retrieving token using password flow failed', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    sinon.stub(publicApplication, 'acquireTokenByUsernamePassword').callsFake(_ => Promise.reject({
      errorCode: 'error',
      errorMessage: `An error has occurred`
    }));
    auth.connection.authType = AuthType.Password;

    auth.ensureAccessToken(resource, logger).then(() => {
      done('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(err.errorMessage, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses browser auth and retrieves a successful response', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    const acquireTokenByCodeStub = sinon.stub(publicApplication, 'acquireTokenByCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));
    auth.connection.authType = AuthType.Browser;

    auth.ensureAccessToken(resource, logger, false).then(() => {
      try {
        assert(acquireTokenByCodeStub.called);
        const args = acquireTokenByCodeStub.args[0][0];
        assert.strictEqual(args.code, httpServerResponse.code, 'Incorrect code');
        assert.strictEqual(args.redirectUri, httpServerResponse.redirectUri, 'Incorrect redirectUri');
        assert.deepStrictEqual(args.scopes, ['https://contoso.sharepoint.com/.default'], 'Incorrect scopes');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('uses browser auth and retrieves a successful response (debug)', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    const acquireTokenByCodeStub = sinon.stub(publicApplication, 'acquireTokenByCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));
    auth.connection.authType = AuthType.Browser;

    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert(acquireTokenByCodeStub.called);
        const args = acquireTokenByCodeStub.args[0][0];
        assert.strictEqual(args.code, httpServerResponse.code, 'Incorrect code');
        assert.strictEqual(args.redirectUri, httpServerResponse.redirectUri, 'Incorrect redirectUri');
        assert.deepStrictEqual(args.scopes, ['https://contoso.sharepoint.com/.default'], 'Incorrect scopes');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('uses browser auth and retrieves an unsuccessful response', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    const acquireTokenByCodeStub = sinon.stub(publicApplication, 'acquireTokenByCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));

    const error = <InteractiveAuthorizationErrorResponse>{
      error: "shortError",
      errorDescription: "errorHasOccurred"
    };

    initializeServerStub.restore();
    initializeServerStub = sinon.stub((auth as any)._authServer, 'initializeServer').callsFake(((connection: Connection, resource: string, resolve: (error: InteractiveAuthorizationCodeResponse) => void, reject: (error: InteractiveAuthorizationErrorResponse) => void) => {
      reject(error);
    }) as any);
    auth.connection.authType = AuthType.Browser;

    auth.ensureAccessToken(resource, logger, false).then(() => {
      done("Should not be called");
    }, (err) => {
      try {
        assert(acquireTokenByCodeStub.notCalled, 'acquireTokenByCodeStub called');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(error));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses browser auth and retrieves an unsuccessful response (debug)', (done) => {
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    const acquireTokenByCodeStub = sinon.stub(publicApplication, 'acquireTokenByCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));

    const error = <InteractiveAuthorizationErrorResponse>{
      error: "shortError",
      errorDescription: "errorHasOccurred"
    };

    initializeServerStub.restore();
    initializeServerStub = sinon.stub((auth as any)._authServer, 'initializeServer').callsFake(((connection: Connection, resource: string, resolve: (error: InteractiveAuthorizationCodeResponse) => void, reject: (error: InteractiveAuthorizationErrorResponse) => void) => {
      reject(error);
    }) as any);
    auth.connection.authType = AuthType.Browser;

    auth.ensureAccessToken(resource, logger, true).then(() => {
      done("Should not be called");
    }, (err) => {
      try {
        assert(acquireTokenByCodeStub.notCalled);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(error));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses browser auth and retrieves retrieving a token but acquisition failed', (done) => {
    const error = {
      errorCode: 'error',
      errorMessage: `An error has occurred`
    };
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    sinon.stub(publicApplication, 'acquireTokenByCode').callsFake(_ => Promise.reject(error));
    auth.connection.authType = AuthType.Browser;

    auth.ensureAccessToken(resource, logger).then(() => {
      assert.fail('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(error));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses browser auth and retrieves retrieving a token but acquisition failed (debug)', (done) => {
    const error = {
      errorCode: 'error',
      errorMessage: `An error has occurred`
    };
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    sinon.stub(publicApplication, 'acquireTokenByCode').callsFake(_ => Promise.reject(error));
    auth.connection.authType = AuthType.Browser;

    auth.ensureAccessToken(resource, logger, true).then(() => {
      assert.fail('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(error));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves token using certificate flow when authType certificate specified', (done) => {
    auth.connection.certificate = base64EncodedPemCert;
    auth.connection.authType = AuthType.Certificate;

    let actualThumbprint: string = '';
    let acquireTokenByClientCredentialStub: any;
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      acquireTokenByClientCredentialStub = sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, false).then(_ => {
      try {
        assert(acquireTokenByClientCredentialStub.called);
        assert.strictEqual(actualThumbprint, "ccf4f2a3c3d209c512b3724bb883a5474c0921dc");
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using certificate flow when authType certificate specified (debug)', (done) => {
    auth.connection.certificate = base64EncodedPemCert;
    auth.connection.authType = AuthType.Certificate;

    let actualThumbprint: string = '';
    let acquireTokenByClientCredentialStub: any;
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      acquireTokenByClientCredentialStub = sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(_ => {
      try {
        assert(acquireTokenByClientCredentialStub.called);
        assert.strictEqual(actualThumbprint, "ccf4f2a3c3d209c512b3724bb883a5474c0921dc");
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using certificate flow when authType certificate and thumbprint specified ', (done) => {
    auth.connection.certificate = base64EncodedPemCert;
    auth.connection.thumbprint = 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc';
    auth.connection.authType = AuthType.Certificate;

    let actualThumbprint: string = '';
    let acquireTokenByClientCredentialStub: any;
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      acquireTokenByClientCredentialStub = sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, false).then(_ => {
      try {
        assert(acquireTokenByClientCredentialStub.called);
        assert.strictEqual(actualThumbprint, "ccf4f2a3c3d209c512b3724bb883a5474c0921dc");
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using certificate flow when authType certificate and thumbprint specified (debug)', (done) => {
    auth.connection.certificate = base64EncodedPemCert;
    auth.connection.thumbprint = 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc';
    auth.connection.authType = AuthType.Certificate;

    let actualThumbprint: string = '';
    let acquireTokenByClientCredentialStub: any;
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      acquireTokenByClientCredentialStub = sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(_ => {
      try {
        assert(acquireTokenByClientCredentialStub.called);
        assert.strictEqual(actualThumbprint, "ccf4f2a3c3d209c512b3724bb883a5474c0921dc");
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using PFX certificate flow when authType certificate specified', (done) => {
    auth.connection.password = 'pass@word1';
    auth.connection.certificate = base64EncodedPfxCert;
    auth.connection.authType = AuthType.Certificate;

    readFileSyncStub.restore();
    let actualThumbprint: string = '';
    let actualCert: string = '';
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      actualCert = cert as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, false).then(_ => {
      try {
        assert.notStrictEqual(actualCert.indexOf('MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQ'), -1);
        assert.strictEqual(actualThumbprint, 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using PFX certificate flow when authType certificate specified (debug)', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.password = 'pass@word1';
    auth.connection.certificate = base64EncodedPfxCert;

    readFileSyncStub.restore();
    let actualThumbprint: string = '';
    let actualCert: string = '';
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      actualCert = cert as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(_ => {
      try {
        assert.notStrictEqual(actualCert.indexOf('MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQ'), -1);
        assert.strictEqual(actualThumbprint, 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using certificate flow when authType certificate and certificateType specified ', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.certificate = base64EncodedPemCert;
    auth.connection.certificateType = CertificateType.Base64;

    let acquireTokenByClientCredentialStub: any;
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      acquireTokenByClientCredentialStub = sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, false).then(() => {
      try {
        assert(acquireTokenByClientCredentialStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using certificate flow when authType certificate and certificateType specified (debug)', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.certificate = base64EncodedPemCert;
    auth.connection.certificateType = CertificateType.Base64;

    let acquireTokenByClientCredentialStub: any;
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      acquireTokenByClientCredentialStub = sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert(acquireTokenByClientCredentialStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using PFX certificate flow when authType certificate and thumbprint specified', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.password = 'pass@word1';
    auth.connection.certificate = base64EncodedPfxCert;
    auth.connection.thumbprint = 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc';

    let actualThumbprint: string = '';
    let actualCert: string = '';
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      actualCert = cert as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, false).then(_ => {
      try {
        assert.notStrictEqual(actualCert.indexOf('MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQ'), -1);
        assert.strictEqual(actualThumbprint, 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using PFX certificate flow when authType certificate and thumbprint specified (debug)', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.password = 'pass@word1';
    auth.connection.certificate = base64EncodedPfxCert;
    auth.connection.thumbprint = 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc';

    let actualThumbprint: string = '';
    let actualCert: string = '';
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      actualCert = cert as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.notStrictEqual(actualCert.indexOf('MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQ'), -1);
        assert.strictEqual(actualThumbprint, 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using PFX certificate flow when authType certificate and certificateType specified (debug)', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.password = 'pass@word1';
    auth.connection.certificate = base64EncodedPfxCert;
    auth.connection.certificateType = CertificateType.Binary;

    let actualThumbprint: string = '';
    let actualCert: string = '';
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      actualThumbprint = thumbprint as string;
      actualCert = cert as string;
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.notStrictEqual(actualCert.indexOf('MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQ'), -1);
        assert.strictEqual(actualThumbprint, 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when PFX certificate flow when authType certificate specified (debug)', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.certificate = base64EncodedPfxCert;
    auth.connection.password = 'abc';

    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(() => {
      done('Expected error');
    }, (err) => {
      try {
        assert.strictEqual(err.toString(), 'Error: PKCS#12 MAC could not be verified. Invalid password?');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving token using certificate flow failed', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.certificate = base64EncodedPemCert;

    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.reject('An error has occurred'));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger).then(() => {
      done('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs error when retrieving token using certificate flow failed in debug mode', (done) => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.certificate = base64EncodedPemCert;

    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert) => {
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, thumbprint, cert);
      sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.reject({ errorCode: 'error', errorMessage: 'An error has occurred' }));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger, true).then(() => {
      done('Got access token');
    }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify({ errorCode: 'error', errorMessage: 'An error has occurred' }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls api with correct params using system managed identity flow when authType identity and Azure VM api', (done) => {
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = undefined;
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://169.254.169.254/metadata/identity/oauth2/token?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2018-02-01');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('gets token using system managed identity flow when authType identity and Azure VM api', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = undefined;
    auth.ensureAccessToken(resource, logger, true).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, 'eyJ0eXAiOiJKV1QiLCJ...');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('calls api with correct params user-assigned managed identity flow when authType identity and client_id and Azure VM api', (done) => {
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://169.254.169.254/metadata/identity/oauth2/token?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2018-02-01&client_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('calls api with correct params user-assigned managed identity flow when authType identity and principal_id and Azure VM api', (done) => {
    const requestStub = sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "error": "invalid_request", "error_description": "Identity not found" } });
      }

      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://169.254.169.254/metadata/identity/oauth2/token?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2018-02-01&principal_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('retrieves token using user-assigned managed identity flow when authType identity and principal_id and Azure VM api', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "error": "invalid_request", "error_description": "Identity not found" } });
      }

      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, 'eyJ0eXAiOiJKV1QiLCJ...');
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when using user-assigned managed identity flow when authType identity and principal_id and Azure VM api', (done) => {
    const requestStub = sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "error": "invalid_request", "error_description": "Identity not found" } });
      }

      return Promise.reject({ error: { "error": "invalid_request", "error_description": "Identity not found" } });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://169.254.169.254/metadata/identity/oauth2/token?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2018-02-01&principal_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        assert.strictEqual(err.error.error_description, 'Identity not found');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles EACCES error when using user-assigned managed identity flow when authType identity and principal_id and Azure VM api', (done) => {
    const requestStub = sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "error": "invalid_request", "error_description": "Identity not found" } });
      }

      return Promise.reject({ error: { "errno": "EACCES", "code": "EACCES", "syscall": "connect", "address": "169.254.169.254", "port": 80 } });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://169.254.169.254/metadata/identity/oauth2/token?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2018-02-01&principal_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        assert.notStrictEqual(err.indexOf('Error while logging with Managed Identity. Please check if a Managed Identity is assigned to the current Azure resource.'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls api with correct params using system managed identity flow when authType identity and Azure Function api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = undefined;
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('calls api with correct params using system managed identity flow when authType identity and Azure Cloud Shell api', (done) => {
    process.env = {
      IDENTITY_ENDPOINT: 'http://localhost:50342/oauth2/token'
    };
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = undefined;
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://localhost:50342/oauth2/token?resource=https%3A%2F%2Fcontoso.sharepoint.com');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('fails with error when authType identity and Azure Cloud Shell api and IDENTITY_ENDPOINT, but userName option specified', (done) => {
    process.env = {
      IDENTITY_ENDPOINT: 'http://localhost:50342/oauth2/token',
      ACC_CLOUD: 'abc'
    };
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve();
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'abc';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.notStrictEqual(err.indexOf('Azure Cloud Shell does not support user-managed identity. You can execute the command without the --userName option to login with user identity'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls api with correct params using MSI when authType identity and Azure Cloud Shell api', (done) => {
    process.env = {
      MSI_ENDPOINT: 'http://localhost:50342/oauth2/token'
    };
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = undefined;
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://localhost:50342/oauth2/token?resource=https%3A%2F%2Fcontoso.sharepoint.com');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('fails with error when authType identity and Azure Cloud Shell api and MSI_ENDPOINT, but userName option specified', (done) => {
    process.env = {
      MSI_ENDPOINT: 'http://localhost:50342/oauth2/token',
      ACC_CLOUD: 'abc'
    };
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve();
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'abc';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.notStrictEqual(err.indexOf('Azure Cloud Shell does not support user-managed identity. You can execute the command without the --userName option to login with user identity'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when using system managed identity flow when authType identity and Azure Function api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({ error: { "StatusCode": 400, "Message": "No Managed Identity found for specified ClientId/ResourceId/PrincipalId.", "CorrelationId": "0507ee4d-c15f-421a-b96b-e71e351bc69a" } });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = undefined;
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        assert.notStrictEqual(err.error.Message.indexOf('No Managed Identity found'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls api with correct params using user-assigned managed identity flow when authType identity and client_id and Azure Functions api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01&client_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('calls api with correct params using user-assigned managed identity flow when authType identity and principal_id and Azure Functions api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "StatusCode": 400, "Message": "No Managed Identity found for specified ClientId/ResourceId/PrincipalId.", "CorrelationId": "0507ee4d-c15f-421a-b96b-e71e351bc69a" } });
      }

      return Promise.resolve({ "access_token": "eyJ0eXA", "expires_on": "1587849030", "resource": "https://contoso.sharepoint.com", "token_type": "Bearer", "client_id": "A04566DF-9A65-4E90-AE3D-574572A16423" });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01&principal_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('handles error when using user-assigned managed identity flow when authType identity and principal_id and Azure Functions api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "StatusCode": 400, "Message": "No Managed Identity found for specified ClientId/ResourceId/PrincipalId.", "CorrelationId": "0507ee4d-c15f-421a-b96b-e71e351bc69a" } });
      }

      return Promise.reject({ error: { "StatusCode": 400, "Message": "No Managed Identity found for specified ClientId/ResourceId/PrincipalId.", "CorrelationId": "0507ee4d-c15f-421a-b96b-e71e351bc69a" } });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01&principal_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        assert.notStrictEqual(err.error.Message.indexOf('No Managed Identity found'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles EACCES error when using user-assigned managed identity flow when authType identity and principal_id and Azure Functions api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "StatusCode": 400, "Message": "No Managed Identity found for specified ClientId/ResourceId/PrincipalId.", "CorrelationId": "0507ee4d-c15f-421a-b96b-e71e351bc69a" } });
      }

      return Promise.reject({ error: { "errno": "EACCES", "code": "EACCES", "syscall": "connect", "address": "169.254.169.254", "port": 80 } });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, false).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01&principal_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        assert.notStrictEqual(err.indexOf('Error while logging with Managed Identity. Please check if a Managed Identity is assigned to the current Azure resource.'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles undefined error when using user-assigned managed identity flow when authType identity and client_id and Azure Functions api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({ error: { "error": "invalid_request", "error_description": "Undefined" } });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01&client_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        assert.notStrictEqual(err.error.error_description.indexOf('Undefined'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles undefined error when using user-assigned managed identity flow when authType identity and principal_id and Azure Functions api', (done) => {
    process.env.IDENTITY_ENDPOINT = 'http://127.0.0.1:41932/MSI/token/';
    process.env.IDENTITY_HEADER = 'AFBA957766234A0CA9F3B6FA3D9582C7';
    const requestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('&client_id=') !== -1) {

        return Promise.reject({ error: { "StatusCode": 400, "Message": "No Managed Identity found for specified ClientId/ResourceId/PrincipalId.", "CorrelationId": "0507ee4d-c15f-421a-b96b-e71e351bc69a" } });
      }
      return Promise.reject({ error: { "error": "Undefined" } });
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = 'a04566df-9a65-4e90-ae3d-574572a16423';
    auth.ensureAccessToken(resource, logger, true).then(() => {
      done(new Error('something is wrong'));
    }, (err) => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01&principal_id=a04566df-9a65-4e90-ae3d-574572a16423');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        assert.notStrictEqual(err.error.error.indexOf('Undefined'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls api with correct params using system managed identity flow when authType identity and Azure Function api using the old env variables', (done) => {
    process.env = {
      MSI_ENDPOINT: 'http://127.0.0.1:41932/MSI/token/',
      MSI_SECRET: 'AFBA957766234A0CA9F3B6FA3D9582C7'
    };
    const requestStub = sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve(JSON.stringify({
        "access_token": "eyJ0eXAiOiJKV1QiLCJ...",
        "client_id": "a04566df-9a65-4e90-ae3d-574572a16423",
        "expires_in": "86399",
        "expires_on": "1587847593",
        "ext_expires_in": "86399",
        "not_before": "1587760893",
        "resource": "https://contoso.sharepoint.com/",
        "token_type": "Bearer"
      }));
    });

    auth.connection.authType = AuthType.Identity;
    auth.connection.userName = undefined;
    auth.ensureAccessToken(resource, logger, true).then(() => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'http://127.0.0.1:41932/MSI/token/?resource=https%3A%2F%2Fcontoso.sharepoint.com&api-version=2019-08-01');
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers.Metadata, true);
        assert.strictEqual((requestStub.lastCall.args[0] as any).headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('returns access token if persisting connection fails', (done) => {
    sinonUtil.restore(auth.storeConnectionInfo);
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    sinon.stub(publicApplication, 'acquireTokenByDeviceCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));

    auth.ensureAccessToken(resource, logger).then((accessToken) => {
      assert.strictEqual(accessToken, 'acc');
      done();
    }, (err) => {
      done(err);
    });
  });

  it('logs error message if persisting connection fails in debug mode', (done) => {
    sinonUtil.restore(auth.storeConnectionInfo);
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    sinon.stub(publicApplication, 'acquireTokenByDeviceCode').callsFake(_ => Promise.resolve({
      expiresOn: new Date(),
      accessToken: 'acc'
    } as any));

    auth.ensureAccessToken(resource, logger, true).then((accessToken) => {
      try {
        assert.strictEqual(accessToken, 'acc');
        assert(loggerLogToStderrSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('configures FileTokenStorage as active connection storage', (done) => {
    const actual = auth.getConnectionStorage();
    try {
      assert(actual instanceof FileTokenStorage);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('configures FileTokenStorage as connections storage', (done) => {
    const actual = auth.getAllConnectionsStorage();
    try {
      assert(actual instanceof FileTokenStorage);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('configures MSAL cache storage as token storage', (done) => {
    const actual = (auth as any).getMsalCacheStorage();
    try {
      assert(actual instanceof FileTokenStorage);
      assert.strictEqual((actual as any).filePath, FileTokenStorage.msalCacheFilePath());
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('restores authentication', (done) => {
    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.strictEqual(auth.connection.active, true);
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err) => {
        done(err);
      });
  });

  it(`doesn't restore authentication when already restored`, (done) => {
    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
    const getConnectionInfoFromStorageStub = sinon.stub(auth as any, 'getConnectionInfoFromStorage').callsFake(() => Promise.resolve());
    auth.connection.active = true;

    auth
      .restoreAuth()
      .then(() => {
        try {
          assert(getConnectionInfoFromStorageStub.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('handles error when restoring authentication', (done) => {
    auth.connection.active = false;
    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
    sinon.stub(auth as any, 'getConnectionInfoFromStorage').callsFake(() => Promise.reject('An error has occurred'));

    auth
      .restoreAuth()
      .then(() => {
        try {
          assert.strictEqual(auth.connection.active, false);
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err) => {
        done(err);
      });
  });

  it('doesn\'t fail when restoring authentication from an incorrect JSON string', (done) => {
    auth.connection.active = false;
    const mockStorage = {
      get: () => Promise.resolve('abc')
    };
    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
    sinon.stub(auth, 'getConnectionStorage').callsFake(() => mockStorage as any);

    auth
      .restoreAuth()
      .then(() => {
        assert.strictEqual(auth.connection.active, false);
        done();
      }, (err) => {
        done(err);
      });
  });

  it('doesn\'t fail when restoring authentication failed', (done) => {
    auth.connection.active = false;
    const mockStorage = {
      get: () => Promise.reject('abc')
    };
    sinonUtil.restore((auth as any).getConnectionInfoFromStorage);
    sinon.stub(auth, 'getConnectionStorage').callsFake(() => mockStorage as any);

    auth
      .restoreAuth()
      .then(() => {
        assert.strictEqual(auth.connection.active, false);
        done();
      }, (err) => {
        done(err);
      });
  });

  it('stores connection information in the configured token storage', async () => {
    sinonUtil.restore(auth.storeConnectionInfo);
    const mockStorage1 = new MockTokenStorage();
    const mockStorage2 = new MockTokenStorage();
    const mockStorageSetStub1 = sinon.stub(mockStorage1, 'set').resolves();
    const mockStorageSetStub2 = sinon.stub(mockStorage2, 'set').resolves();
    sinon.stub(auth, 'getConnectionStorage').returns(mockStorage1);
    sinon.stub(auth, 'getAllConnectionsStorage').returns(mockStorage2);

    await auth.storeConnectionInfo();
    assert(mockStorageSetStub1.called, 'Active connection Storage not set');
    assert(mockStorageSetStub2.called, 'All connections Storage not set');
  });

  it('clears connection information in the configured token storage', async () => {
    const mockStorage1 = new MockTokenStorage();
    const mockStorage2 = new MockTokenStorage();
    const mockStorage3 = new MockTokenStorage();
    const mockStorageRemoveStub1 = sinon.stub(mockStorage1, 'remove').resolves();
    const mockStorageRemoveStub2 = sinon.stub(mockStorage2, 'remove').resolves();
    const mockStorageRemoveStub3 = sinon.stub(mockStorage3, 'remove').resolves();
    sinon.stub(auth, 'getConnectionStorage').returns(mockStorage1);
    sinon.stub(auth, 'getAllConnectionsStorage').returns(mockStorage2);
    sinon.stub(auth as any, 'getMsalCacheStorage').callsFake(() => mockStorage3);

    await auth.clearConnectionInfo();
    assert(mockStorageRemoveStub1.called, 'Active connection Storage not cleared');
    assert(mockStorageRemoveStub2.called, 'All connections Storage not cleared');
    assert(mockStorageRemoveStub3.called, 'token storage or MSAL cache not cleared');
  });

  it('removes a connection from the configured token storage', async () => {
    sinonUtil.restore(auth.removeConnectionInfo);
    const mockStorage1 = new MockTokenStorage();
    const mockStorage2 = new MockTokenStorage();
    const mockStorageSetStub1 = sinon.stub(mockStorage1, 'remove').resolves();
    const mockStorageSetStub2 = sinon.stub(mockStorage2, 'set').resolves();
    sinon.stub(auth as any, 'getPublicClient').callsFake(_ => publicApplication);
    sinon.stub(auth, 'getConnectionStorage').returns(mockStorage1);
    sinon.stub(auth, 'getAllConnectionsStorage').returns(mockStorage2);

    await auth.removeConnectionInfo(auth.connection, logger, false);
    assert(mockStorageSetStub1.called, 'Active connection Storage not removed');
    assert(mockStorageSetStub2.called, 'All connections Storage not cleared');
  });

  it('resets connection information on logout', () => {
    auth.connection.active = true;
    auth.connection.accessTokens[resource] = {
      expiresOn: new Date().toISOString(),
      accessToken: 'abc'
    };
    auth.connection.authType = AuthType.Certificate;
    auth.connection.userName = 'user';
    auth.connection.password = 'pwd';
    auth.connection.certificateType = CertificateType.Binary;
    auth.connection.certificate = 'cert';
    auth.connection.thumbprint = 'thumb';
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    auth.connection.spoTenantId = '123';

    auth.connection.deactivate();

    assert.strictEqual(auth.connection.active, false, 'connected');
    assert.strictEqual(JSON.stringify(auth.connection.accessTokens), JSON.stringify({}), 'accessTokens');
    assert.strictEqual(auth.connection.authType, AuthType.DeviceCode, 'authType');
    assert.strictEqual(auth.connection.userName, undefined, 'userName');
    assert.strictEqual(auth.connection.password, undefined, 'password');
    assert.strictEqual(auth.connection.certificateType, CertificateType.Unknown, 'certificateType');
    assert.strictEqual(auth.connection.certificate, undefined, 'certificate');
    assert.strictEqual(auth.connection.thumbprint, undefined, 'thumbprint');
    assert.strictEqual(auth.connection.spoUrl, undefined, 'spoUrl');
    assert.strictEqual(auth.connection.spoTenantId, undefined, 'tenantId');
  });

  it('uses the Microsoft Graph to authenticate', () => {
    assert.strictEqual(auth.defaultResource, 'https://graph.microsoft.com');
  });

  it('correctly retrieves resource from the root SharePoint site URL without trailing slash', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com'), 'https://contoso.sharepoint.com');
  });

  it('correctly retrieves resource from the root SharePoint site URL with trailing slash', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com/'), 'https://contoso.sharepoint.com');
  });

  it('correctly retrieves resource from a SharePoint subsite', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com/subsite'), 'https://contoso.sharepoint.com');
  });

  it('correctly retrieves resource from a SharePoint site collection', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://contoso.sharepoint.com/sites/team-a'), 'https://contoso.sharepoint.com');
  });

  it('correctly retrieves resource for https://api.powerapps.com', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://api.powerapps.com'), 'https://service.powerapps.com/');
  });

  it('correctly retrieves resource for https://api.bap.microsoft.com', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://api.bap.microsoft.com'), 'https://service.powerapps.com/');
  });

  it('correctly retrieves resource for https://api.powerbi.com', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://api.powerbi.com'), 'https://analysis.windows.net/powerbi/api');
  });

  it('correctly retrieves resource for https://api.flow.microsoft.com', () => {
    assert.strictEqual(Auth.getResourceFromUrl('https://api.flow.microsoft.com'), 'https://management.azure.com/');
  });

  it('returns undefined if access token is not set when determining auth type', () => {
    assert.strictEqual(accessToken.isAppOnlyAccessToken(''), undefined);
  });

  it(`returns undefined if access token is not valid`, () => {
    assert.strictEqual(accessToken.isAppOnlyAccessToken('123.456'), undefined);
  });

  it('returns public client for device code auth', async () => {
    auth.connection.authType = AuthType.DeviceCode;
    const actualClientApp = await (auth as any).getPublicClient(logger, false);
    assert(actualClientApp instanceof msal.PublicClientApplication);
  });

  it('returns public client for password auth', async () => {
    auth.connection.authType = AuthType.Password;
    const actualClientApp = await (auth as any).getPublicClient(logger, false);
    assert(actualClientApp instanceof msal.PublicClientApplication);
  });

  it('changes tenant for a multitenant app for password auth to organizations', async () => {
    auth.connection.authType = AuthType.Password;
    auth.connection.tenant = 'common';
    await (auth as any).getPublicClient(logger, false);
    assert.strictEqual(auth.connection.tenant, 'organizations');
  });

  it('returns public client for browser auth', async () => {
    auth.connection.authType = AuthType.Browser;
    const actualClientApp = await (auth as any).getPublicClient(logger, false);
    assert(actualClientApp instanceof msal.PublicClientApplication);
  });

  it('returns confidential client for certificate auth', async () => {
    auth.connection.authType = AuthType.Certificate;
    auth.connection.thumbprint = 'ccf4f2a3c3d209c512b3724bb883a5474c0921dc';
    const actualClientApp = await (auth as any).getConfidentialClient(logger, false, auth.connection.thumbprint as string, auth.connection.password, undefined);
    assert(actualClientApp instanceof msal.ConfidentialClientApplication);
  });

  it('returns confidential client for secret auth', async () => {
    auth.connection.authType = AuthType.Secret;
    auth.connection.secret = 'sOmeToPsecRetValue';
    const actualClientApp = await (auth as any).getConfidentialClient(logger, false, undefined, undefined, auth.connection.secret);
    assert(actualClientApp instanceof msal.ConfidentialClientApplication);
  });

  it('retrieves token using client secret flow when authType "secret" specified', (done) => {
    auth.connection.authType = AuthType.Secret;
    auth.connection.secret = "SomeSecretValue";

    let acquireTokenByClientCredentialStub: any;
    let originalGetConfidentialClient = (auth as any).getConfidentialClient;
    originalGetConfidentialClient = originalGetConfidentialClient.bind(auth);
    sinon.stub(auth as any, 'getConfidentialClient').callsFake(async (logger, debug, thumbprint, cert, clientSecret) => {
      const confidentialApplication = await originalGetConfidentialClient(logger, debug, undefined, undefined, clientSecret);
      acquireTokenByClientCredentialStub = sinon.stub(confidentialApplication, 'acquireTokenByClientCredential').callsFake(_ => Promise.resolve({
        expiresOn: new Date(),
        accessToken: 'acc'
      } as any));
      return confidentialApplication;
    });
    sinon.stub(tokenCache, 'getAllAccounts').resolves([]);

    auth.ensureAccessToken(resource, logger).then(() => {
      try {
        assert(acquireTokenByClientCredentialStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    }, (err) => {
      done(err);
    });
  });

  it('configures cloud for auth to AzureChina for China cloud', async () => {
    auth.connection.cloudType = CloudType.China;
    const actual: msal.Configuration = await (auth as any).getAuthClientConfiguration(logger, false);
    assert.strictEqual(actual.auth.azureCloudOptions?.azureCloudInstance, msal.AzureCloudInstance.AzureChina);
  });

  it('configures cloud for auth to AzureUsGovernment for USGov cloud', async () => {
    auth.connection.cloudType = CloudType.USGov;
    const actual: msal.Configuration = await (auth as any).getAuthClientConfiguration(logger, false);
    assert.strictEqual(actual.auth.azureCloudOptions?.azureCloudInstance, msal.AzureCloudInstance.AzureUsGovernment);
  });

  it('configures cloud for auth to AzureUsGovernment for USGovHigh cloud', async () => {
    auth.connection.cloudType = CloudType.USGovHigh;
    const actual: msal.Configuration = await (auth as any).getAuthClientConfiguration(logger, false);
    assert.strictEqual(actual.auth.azureCloudOptions?.azureCloudInstance, msal.AzureCloudInstance.AzureUsGovernment);
  });

  it('configures cloud for auth to AzureUsGovernment for USGovDoD cloud', async () => {
    auth.connection.cloudType = CloudType.USGovDoD;
    const actual: msal.Configuration = await (auth as any).getAuthClientConfiguration(logger, false);
    assert.strictEqual(actual.auth.azureCloudOptions?.azureCloudInstance, msal.AzureCloudInstance.AzureUsGovernment);
  });

  it(`loads all connections from storage if not already loaded`, async () => {
    sinonUtil.restore((auth as any).getAllConnectionsFromStorage);
    const mockStorage = {
      get: () => { return '[{ "name": "abc" }]'; }
    };
    const stub = sinon.stub(auth, 'getAllConnectionsStorage').callsFake(() => mockStorage as any);

    auth.connection.active = true;
    (auth as any)._allConnections = undefined;

    const connections = await auth.getAllConnections();
    assert(stub.called);
    assert.strictEqual(connections.length, 1);
  });

  it(`handles failure while loading all connections from storage if not already loaded`, async () => {
    sinonUtil.restore((auth as any).getAllConnectionsFromStorage);
    const mockStorage = {
      get: () => Promise.reject('Invalid!')
    };
    const stub = sinon.stub(auth, 'getAllConnectionsStorage').callsFake(() => mockStorage as any);

    auth.connection.active = true;
    (auth as any)._allConnections = undefined;

    const connections = await auth.getAllConnections();
    assert(stub.called);
    assert.strictEqual(connections.length, 0);
  });
});
