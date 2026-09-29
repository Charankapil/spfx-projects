import * as React from 'react';
import { Checkbox, Dropdown, IDropdownOption, Stack, Toggle } from '@fluentui/react';

import { IRestoreOptions } from '../models/IRestoreOptions';
import styles from './ReInherit.module.scss';

interface IOptionsPanelProps {
  options: IRestoreOptions;
  onChange: (options: IRestoreOptions) => void;
  disabled?: boolean;
}

const DEPTH_OPTIONS: IDropdownOption[] = [
  { key: 0, text: 'All levels (down to the last file)' },
  { key: 1, text: '1 level - top-level folders and files' },
  { key: 2, text: '2 levels' },
  { key: 3, text: '3 levels' },
  { key: 4, text: '4 levels' },
  { key: 5, text: '5 levels' },
  { key: 10, text: '10 levels' }
];

export const OptionsPanel: React.FC<IOptionsPanelProps> = ({ options, onChange, disabled }) => {
  const set = (patch: Partial<IRestoreOptions>): void => onChange({ ...options, ...patch });
  const nothingBelow =
    options.recursive &&
    !options.includeSubsites &&
    !options.includeLists &&
    !options.includeFolders &&
    !options.includeFiles;

  return (
    <Stack tokens={{ childrenGap: 14 }}>
      <Toggle
        label="Include everything below the selection"
        inlineLabel={false}
        onText="Recursive - subsites, lists, folders and files below"
        offText="Only the selected sites, lists and folders"
        checked={options.recursive}
        disabled={disabled}
        onChange={(_, checked) => set({ recursive: !!checked })}
      />

      <Checkbox
        label="Restore the selected objects themselves"
        checked={options.includeSelected}
        disabled={disabled}
        onChange={(_, checked) => set({ includeSelected: !!checked })}
      />

      {options.recursive && (
        <div className={styles.optionGroup}>
          <div className={styles.optionGroupTitle}>Below the selection, restore</div>
          <Stack tokens={{ childrenGap: 8 }}>
            <Checkbox
              label="Subsites (and walk into them)"
              checked={options.includeSubsites}
              disabled={disabled}
              onChange={(_, checked) => set({ includeSubsites: !!checked })}
            />
            <Checkbox
              label="Lists and libraries"
              checked={options.includeLists}
              disabled={disabled}
              onChange={(_, checked) => set({ includeLists: !!checked })}
            />
            <Checkbox
              label="Folders"
              checked={options.includeFolders}
              disabled={disabled}
              onChange={(_, checked) => set({ includeFolders: !!checked })}
            />
            <Checkbox
              label="Files and list items"
              checked={options.includeFiles}
              disabled={disabled}
              onChange={(_, checked) => set({ includeFiles: !!checked })}
            />
          </Stack>
          <Dropdown
            className={styles.depthDropdown}
            label="How deep into folders"
            selectedKey={options.maxDepth}
            options={DEPTH_OPTIONS}
            disabled={disabled || (!options.includeFolders && !options.includeFiles)}
            onChange={(_, option) => option && set({ maxDepth: option.key as number })}
          />
          <div className={styles.hint}>
            Measured from the selected library or folder. For example, 1 level restores a library&apos;s
            top-level folders and files but leaves anything inside those folders unique.
          </div>
          {nothingBelow && (
            <div className={styles.warningText}>Nothing below the selection is ticked, so only the selected objects are checked.</div>
          )}
        </div>
      )}

      <Toggle
        label="Back up current permissions into the report"
        onText="On - who had access is recorded before each reset"
        offText="Off - faster, but the report won't show who had access"
        checked={options.backupPermissions}
        disabled={disabled}
        onChange={(_, checked) => set({ backupPermissions: !!checked })}
      />
    </Stack>
  );
};
