export interface PlannerTasksValue {
  planId: string;
  bucketId: string;
  title: string;
  orderHint: string;
  assigneePriority: string;
  percentComplete: number;
  startDateTime: Date;
  createdDateTime: Date;
  dueDateTime: Date;
  hasDescription: boolean;
  previewType: string;
  completedDateTime?: Date;
  completedBy: CompletedBy;
  referenceCount: number;
  checklistItemCount: number;
  activeChecklistItemCount: number;
  conversationThreadId?: any;
  id: string;
  createdBy: CreatedBy;
  appliedCategories: AppliedCategories;
  assignments: Object[];
}

export interface CreatedBy {
  user: User;
}

export interface AppliedCategories {}

export interface User {
  displayName?: any;
  id: string;
}

export interface CompletedBy {
  user: User;
}
