import {
  IActivity,
  ICompletedPayload,
  IDelegatedPayload,
  IRecipient,
  IRequest,
  RequestState
} from '../models/types';

/**
 * Derives a request's live state from the sender's request item and the
 * activity items. Pure functions only - no SharePoint calls - so the rules
 * that decide who may sign, and whether a request is complete, live in one
 * place.
 *
 * Every activity is checked against the identity SharePoint stamped on it
 * (Author), never against anything written in the payload. An activity that
 * was edited after it was created is ignored and reported, because the
 * Activity list only lets people edit their own items and a signature must
 * not change once given.
 */

export type SlotStatus = 'waiting' | 'pending' | 'signed' | 'declined';

export interface ISlotState {
  /** The recipient as originally added by the sender. */
  original: IRecipient;
  /** Who currently holds the slot, after any delegation. */
  current: IRecipient;
  status: SlotStatus;
  viewedOn?: string;
  signedOn?: string;
  signedActivity?: IActivity;
  declinedActivity?: IActivity;
  delegations: IActivity[];
}

export interface IRequestStatus {
  state: RequestState;
  slots: ISlotState[];
  /** Signer slots whose turn it is right now. */
  current: ISlotState[];
  ccs: IRecipient[];
  completed?: IActivity;
  cancelled?: IActivity;
  declined?: IActivity;
  /** Most recent time recipients were notified (sent or reminded). */
  lastNotifiedOn: string;
  /** Activities that were ignored, with the reason - shown in the audit trail. */
  warnings: string[];
  /** Why each ignored activity (by id) did not count. */
  ignored: { [activityId: number]: string };
  signedCount: number;
  signerCount: number;
}

/** SharePoint stamps Created and Modified together; allow a little clock slack. */
const EDIT_TOLERANCE_MS = 5000;

export function sameEmail(a: string | undefined, b: string | undefined): boolean {
  return !!a && !!b && a.trim().toLowerCase() === b.trim().toLowerCase();
}

export function wasEdited(activity: IActivity): boolean {
  return new Date(activity.modified).getTime() - new Date(activity.created).getTime() > EDIT_TOLERANCE_MS;
}

/** "YYYY-MM-DD" as the last millisecond of that day in local time (new Date('YYYY-MM-DD') would be UTC). */
export function endOfDay(isoDate: string): Date {
  const [y, m, d] = isoDate.substring(0, 10).split('-').map(n => parseInt(n, 10));
  return new Date(y, m - 1, d, 23, 59, 59, 999);
}

export function toIsoDate(date: Date): string {
  const pad = (n: number): string => (n < 10 ? '0' : '') + n;
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/** End of the expiry day in local time. */
export function isExpired(expiresOn: string | undefined, now: Date): boolean {
  if (!expiresOn) {
    return false;
  }
  return now.getTime() > endOfDay(expiresOn).getTime();
}

export function computeStatus(request: IRequest, activities: IActivity[], now: Date = new Date()): IRequestStatus {
  const payload = request.payload;
  const signers = payload.recipients.filter(r => r.role === 'signer');
  const warnings: string[] = [];
  const ignored: { [activityId: number]: string } = {};
  const ignore = (a: IActivity, reason: string): void => {
    ignored[a.id] = reason;
    warnings.push(`Ignored a "${a.type}" entry by ${a.authorName}: ${reason}.`);
  };

  const slots: ISlotState[] = signers.map(r => ({
    original: r,
    current: { ...r },
    status: 'waiting' as SlotStatus,
    delegations: []
  }));
  const slotById: { [id: string]: ISlotState } = {};
  slots.forEach(s => (slotById[s.original.id] = s));

  const isSender = (a: IActivity): boolean => sameEmail(a.authorEmail, request.authorEmail);
  const isParticipant = (a: IActivity): boolean =>
    isSender(a) || slots.some(s => sameEmail(s.current.email, a.authorEmail) || sameEmail(s.original.email, a.authorEmail));

  let cancelled: IActivity | undefined;
  let declined: IActivity | undefined;
  let completed: IActivity | undefined;
  let lastNotifiedOn = payload.sentOn;

  const ordered = activities
    .filter(a => a.requestId === request.requestId)
    .sort((a, b) => new Date(a.created).getTime() - new Date(b.created).getTime() || a.id - b.id);

  for (const a of ordered) {
    if (wasEdited(a)) {
      ignore(a, 'it was changed after it was recorded');
      continue;
    }
    const slot = a.slot ? slotById[a.slot] : undefined;
    const ownsSlot = !!slot && sameEmail(slot.current.email, a.authorEmail);

    switch (a.type) {
      case 'sent':
      case 'reminded':
        if (isSender(a) && new Date(a.created).getTime() > new Date(lastNotifiedOn).getTime()) {
          lastNotifiedOn = a.created;
        }
        break;
      case 'viewed':
        if (slot && ownsSlot && !slot.viewedOn) {
          slot.viewedOn = a.created;
        }
        break;
      case 'signed':
        if (!slot || !ownsSlot) {
          ignore(a, 'they are not the signer for that slot');
        } else if (slot.status === 'signed' || slot.status === 'declined' || cancelled || declined) {
          // Duplicate or too late - the first decisive action stands.
        } else if (!isTurn(slot, slots)) {
          ignore(a, 'it was given before their turn');
        } else {
          slot.status = 'signed';
          slot.signedOn = a.created;
          slot.signedActivity = a;
        }
        break;
      case 'declined':
        if (slot && ownsSlot && slot.status !== 'signed' && !cancelled && !declined) {
          slot.status = 'declined';
          slot.declinedActivity = a;
          declined = a;
        }
        break;
      case 'delegated': {
        const to = a.payload as IDelegatedPayload;
        if (slot && ownsSlot && slot.status !== 'signed' && slot.status !== 'declined' && to && to.toEmail) {
          slot.delegations.push(a);
          slot.current = { ...slot.current, name: to.toName || to.toEmail, email: to.toEmail };
          slot.viewedOn = undefined;
        }
        break;
      }
      case 'cancelled':
        if (isSender(a) && !cancelled && !completed) {
          cancelled = a;
        }
        break;
      case 'completed':
        if (isParticipant(a) && !completed) {
          completed = a;
        }
        break;
      default:
        break;
    }
  }

  if (!cancelled && request.senderStatus === 'Cancelled') {
    warnings.push('The request is marked cancelled but no cancellation was recorded in the activity log.');
  }

  const allSigned = slots.length > 0 && slots.every(s => s.status === 'signed');
  let state: RequestState;
  if (cancelled || request.senderStatus === 'Cancelled') {
    state = 'Cancelled';
  } else if (declined) {
    state = 'Declined';
  } else if (allSigned) {
    state = 'Completed';
  } else if (isExpired(payload.expiresOn, now)) {
    state = 'Expired';
  } else {
    state = 'InProgress';
  }

  const current = state === 'InProgress' ? slots.filter(s => s.status !== 'signed' && isTurn(s, slots)) : [];
  current.forEach(s => (s.status = 'pending'));

  return {
    state,
    slots,
    current,
    ccs: payload.recipients.filter(r => r.role === 'cc'),
    completed: state === 'Completed' ? completed : undefined,
    cancelled,
    declined,
    lastNotifiedOn,
    warnings,
    ignored,
    signedCount: slots.filter(s => s.status === 'signed').length,
    signerCount: slots.length
  };
}

/** A slot's turn has come once every signer with a lower order number has signed. */
function isTurn(slot: ISlotState, slots: ISlotState[]): boolean {
  return slots.every(s => s.original.order >= slot.original.order || s.status === 'signed');
}

export function slotForUser(status: IRequestStatus, email: string): ISlotState | undefined {
  return status.current.find(s => sameEmail(s.current.email, email));
}

export function completedPayload(status: IRequestStatus): ICompletedPayload | undefined {
  return status.completed ? (status.completed.payload as ICompletedPayload) : undefined;
}

/** Is a reminder due under the request's automatic reminder setting? */
export function reminderDue(request: IRequest, status: IRequestStatus, now: Date = new Date()): boolean {
  const days = request.payload.reminderDays;
  if (!days || status.state !== 'InProgress') {
    return false;
  }
  const last = new Date(status.lastNotifiedOn).getTime();
  return now.getTime() - last >= days * 24 * 60 * 60 * 1000;
}
