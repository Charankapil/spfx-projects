import * as React from 'react';
import { ISiteGrowth } from '../../services/GrowthEngine';
import { Pill } from '../shared/ui';

export function growthSummary(s: { reasons: string[]; status: string }): string {
  return s.reasons.length ? s.reasons.join(' ') : s.status === 'baseline' ? 'Collecting a baseline. Capture again in a day or more.' : 'Normal growth.';
}

export const statusPill = (s: Pick<ISiteGrowth, 'status'>): React.ReactNode =>
  s.status === 'critical' ? <Pill kind="critical">Growing fast</Pill> : s.status === 'warning' ? <Pill kind="warning">Watch</Pill> : s.status === 'baseline' ? <Pill>Baseline</Pill> : <Pill kind="good">Normal</Pill>;
