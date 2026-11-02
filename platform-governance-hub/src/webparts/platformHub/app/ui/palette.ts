import { ITheme } from '@fluentui/react';

/**
 * Categorical chart colours in a fixed, CVD-validated order (light and dark steps
 * of the same hues). Series take slots in order and never cycle; past eight,
 * data is folded into "Other".
 */
const LIGHT = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948'];
const DARK = ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#008300', '#9085e9', '#e66767'];

export function seriesColor(index: number, isDark: boolean): string {
  const list = isDark ? DARK : LIGHT;
  return list[Math.min(index, list.length - 1)];
}

/** "Other" buckets use a neutral so they never compete with real categories. */
export function otherColor(theme: ITheme): string {
  return theme.palette.neutralTertiary;
}

export interface IStatusColors {
  good: string;
  warning: string;
  critical: string;
  info: string;
}

export function statusColors(isDark: boolean): IStatusColors {
  return isDark
    ? { good: '#54b054', warning: '#e0a000', critical: '#f07b7b', info: '#5ea7f0' }
    : { good: '#107c10', warning: '#a35c00', critical: '#c50f1f', info: '#0f6cbd' };
}

export function chartInk(theme: ITheme): { text: string; muted: string; grid: string; surface: string } {
  return {
    text: theme.semanticColors.bodyText,
    muted: theme.semanticColors.bodySubtext,
    grid: theme.palette.neutralLight,
    surface: theme.semanticColors.bodyBackground
  };
}
