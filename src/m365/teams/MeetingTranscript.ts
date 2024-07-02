export interface MeetingTranscript {
  id: string;
  meetingId: string;
  meetingOrganizerId: string;
  transcriptContentUrl: string;
  createdDateTime: Date;
}