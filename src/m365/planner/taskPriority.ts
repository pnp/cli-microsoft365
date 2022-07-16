export const taskPriority = {
  /** Priority labels used in Planner. */
  priorityValues: ["Urgent", "Important", "Medium", "Low"],

  /**
   * Transform priority label to the corresponding number value.
   * @param priority Priority label or number.
   */
  getPriorityValue(priority?: string | number): number | undefined {
    if (typeof priority === "string") {
      switch (priority.toLowerCase()) {
        case "urgent":
          return 1;
        case "important":
          return 3;
        case "medium":
          return 5;
        case "low":
          return 9;
      }
    }
    
    return priority as number | undefined;
  }
};