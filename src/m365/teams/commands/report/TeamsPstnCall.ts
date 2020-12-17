export interface TeamsPstnCall {

  id: string;
  callId: string;
  userId: string;
  userPrincipalName: string;
  userDisplayName: string;
  startDateTime: string;
  endDateTime: string;
  duration: number;
  charge: number;
  callType: string;
  currency: string;
  calleeNumber: string;
  usageCountryCode: string;
  tenantCountryCode: string;
  connectionCharge: number;
  callerNumber: string;
  destinationContext: string;
  destinationName: string;
  conferenceId: string;
  licenseCapability: string;
  inventoryType: string;

}