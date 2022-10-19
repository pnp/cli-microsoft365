export interface EventReceiver {
  ReceiverAssembly: string;
  ReceiverClass: string;
  ReceiverId: string;
  ReceiverName: string;
  SequenceNumber: number;
  Synchronization: number;
  EventType: number;
  ReceiverUrl?: string;
}