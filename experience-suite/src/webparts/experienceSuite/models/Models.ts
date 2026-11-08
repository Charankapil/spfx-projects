import { WebPartContext } from '@microsoft/sp-webpart-base';

export type ModuleKey = 'home' | 'requests' | 'documents' | 'sites' | 'approvals' | 'team' | 'people';
export type TaskStatus = 'Not Started' | 'In Progress' | 'Completed';
export type ApprovalDecision = 'Approved' | 'Rejected';

export interface IExperienceSuiteProps { title: string; context: WebPartContext; enablePeopleSearch: boolean; }
export interface ICurrentUser { Id: number; Title: string; Email: string; LoginName: string; }

export interface IListItem {
  Id: number;
  Title: string;
  Description?: string;
  Status?: string;
  Category?: string;
  Owner?: string;
  Created?: string;
  Modified?: string;
  DueDate?: string;
  RequesterEmail?: string;
  RequesterName?: string;
  ApproverEmail?: string;
  DecisionBy?: string;
  DecisionDate?: string;
  Author?: { Title?: string; Id?: number };
  Editor?: { Title?: string };
  /** Item version tag used for optimistic concurrency (If-Match). */
  etag?: string;
}

export interface IPerson { title: string; email: string; department: string; jobTitle: string; location: string; picture: string; skills: string; profileUrl: string; }
export interface ISearchDocument { title: string; path: string; author: string; modified: string; fileType: string; }

/** Result of loading one backing list; a failure on one list must not blank out the whole suite. */
export interface IListLoadResult { items: IListItem[]; error?: string; }
