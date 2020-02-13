import * as TestResponses from './approleassignment-list-responses.spec';
import * as TestConstants from './approleassignment-list-constants.spec';

export class requestStub {
  static retrieveAppRoles = ((opts: any) => {
    // we need to fake three calls:
    // 2. get the service principal for the assigned resource(s)
    // 3. get the app roles of the resource

    // query for service principal
    if (opts.url.indexOf(`/myorganization/servicePrincipals?api-version=1.6&$expand=appRoleAssignments&$filter=`) > -1) {
      // by app id
      if (opts.url.indexOf(`appId eq '${TestConstants.CommandActionParameters.appIdWithRoleAssignments}'`) > -1) {
        return Promise.resolve(TestResponses.servicePrincipalCollections.ServicePrincipalByAppId);
      }
      // by display name
      if (opts.url.indexOf(`displayName eq '${encodeURIComponent(TestConstants.CommandActionParameters.appNameWithRoleAssignments)}'`) > -1) {
        return Promise.resolve(TestResponses.servicePrincipalCollections.ServicePrincipalByDisplayName);
      }
      // by app id: no app role assignments
      if (opts.url.indexOf(`appId eq '${TestConstants.CommandActionParameters.appIdWithNoRoleAssignments}'`) > -1) {
        return Promise.resolve(TestResponses.servicePrincipalCollections.ServicePrincipalByAppIdNotFound);
      }
      // by app id: does not exist
      if (opts.url.indexOf(`appId eq '${TestConstants.CommandActionParameters.invalidAppId}'`) > -1) {
        return Promise.resolve(TestResponses.servicePrincipalCollections.ServicePrincipalByAppIdNotFound);
      }
    }

    if (opts.url.indexOf(`/myorganization/servicePrincipals/${TestConstants.InternalRequestStub.customAppId}?api-version=1.6`) > -1) {
        return Promise.resolve(TestResponses.servicePrincipalObject.servicePrincipalCustomAppWithAppRole);
    }

    if (opts.url.indexOf(`/myorganization/servicePrincipals/${TestConstants.InternalRequestStub.microsoftGraphAppId}?api-version=1.6`) > -1) {
      return Promise.resolve(TestResponses.servicePrincipalObject.servicePrincipalMicrosoftGraphWithAppRole);
    }

    return Promise.reject('Invalid request');
  })
}