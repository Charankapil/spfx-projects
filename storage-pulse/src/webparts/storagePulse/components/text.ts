import * as strings from 'StoragePulseWebPartStrings';

import { MAX_AGE_MONTHS } from '../models/IScanResult';

/** Fills "{name}" placeholders in a localized string. */
export function format(template: string, values: { [key: string]: string | number }): string {
  return template.replace(/\{(\w+)\}/g, (match, key: string) =>
    Object.prototype.hasOwnProperty.call(values, key) ? String(values[key]) : match
  );
}

export function thresholdLabel(months: number): string {
  if (months % 12 === 0) {
    const years = months / 12;
    return years === 1 ? strings.PeriodOneYear : format(strings.PeriodYears, { count: years });
  }
  return format(strings.PeriodMonths, { count: months });
}

export function describeAge(months: number): string {
  if (months <= 0) {
    return strings.AgeThisMonth;
  }
  if (months >= MAX_AGE_MONTHS) {
    return strings.AgeTenPlus;
  }
  if (months < 12) {
    return months === 1 ? strings.AgeOneMonth : format(strings.AgeMonths, { count: months });
  }
  const years = Math.floor(months / 12);
  const rest = months % 12;
  if (rest === 0) {
    return years === 1 ? strings.AgeOneYear : format(strings.AgeYears, { count: years });
  }
  return years === 1
    ? format(strings.AgeOneYearMonths, { months: rest })
    : format(strings.AgeYearsMonths, { years, months: rest });
}

/** Age band labels for the chart, in AGE_BANDS order. */
export function bandLabels(): string[] {
  return strings.AgeBandLabels.split('|');
}

let categoryLabelCache: { [key: string]: string } | undefined;

export function categoryLabel(key: string): string {
  if (!categoryLabelCache) {
    categoryLabelCache = {};
    for (const pair of strings.CategoryLabels.split('|')) {
      const eq = pair.indexOf('=');
      if (eq > 0) {
        categoryLabelCache[pair.substring(0, eq)] = pair.substring(eq + 1);
      }
    }
  }
  return categoryLabelCache[key] || key;
}
