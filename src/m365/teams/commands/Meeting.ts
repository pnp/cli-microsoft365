export interface Meeting {
  id: string;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  changeKey: string;
  categories: string[];
  transactionId: string;
  originalStartTimeZone: string;
  originalEndTimeZone: string;
  iCalUId: string;
  reminderMinutesBeforeStart: number;
  isReminderOn: boolean;
  hasAttachments: boolean;
  subject: string;
  bodyPreview: string;
  importance: string;
  sensitivity: string;
  isAllDay: boolean;
  isCancelled: boolean;
  isOrganizer: boolean;
  responseRequested: boolean;
  seriesMasterId?: string;
  showAs: string;
  type: string;
  webLink: string;
  onlineMeetingUrl?: string;
  isOnlineMeeting: boolean;
  onlineMeetingProvider: string;
  allowNewTimeProposals: boolean;
  isDraft: boolean;
  hideAttendees: boolean;
  responseStatus: Response;
  body: Body;
  start: MeetingDate;
  end: MeetingDate;
  location: Location;
  locations: Location[];
  attendees: User[];
  organizer: User;
  onlineMeeting: OnlineMeeting;
}

interface Response {
  response: string;
  time: Date;
}

interface Body {
  contentType: string;
  content: string;
}

interface MeetingDate {
  dateTime: Date;
  timeZone: string;
}

interface Location {
  displayName: string;
  locationType: string;
  uniqueIdType: string;
  address: Address;
  coordinates: Coordinates;
}

interface User {
  emailAddress: EmailAddress;
}

interface OnlineMeeting {
  joinUrl: string;
}

interface EmailAddress {
  name: string;
  address: string;
}

interface Address {
  city?: string;
  countryOrRegion?: string;
  postalCode?: string;
  state?: string;
  street?: string;
}

interface Coordinates {
  accuracy: number;
  altitude: number;
  altitudeAccuracy: number;
  latitude: number;
  longitude: number;
}