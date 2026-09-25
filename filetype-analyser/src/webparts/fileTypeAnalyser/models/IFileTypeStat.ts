export interface IFileTypeStat {
  extension: string;
  count: number;
  /**
   * Site-wide storage estimate for this extension, from search size bands
   * (current file versions only). Only set on the site-wide totals.
   */
  estimatedBytes?: number;
}
