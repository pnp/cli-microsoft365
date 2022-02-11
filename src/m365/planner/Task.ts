export interface Task {
  id: string;
  title?: string;
  startDateTime?: Date;
  dueDateTime: Date;
  completedDateTime?: Date;
  planId?: string;
  bucketId?: string;
}

export interface BetaTask extends Task {
  priority?: number;
}
