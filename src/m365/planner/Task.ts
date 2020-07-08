export interface Task {
  id: string;
  title?: string;
  startDateTime?: Date;
  dueDateTime: Date;
  completedDateTime?: Date;
  planId?: string;
  bucketId?: string;
}
