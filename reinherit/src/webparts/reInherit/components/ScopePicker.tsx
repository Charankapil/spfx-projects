import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import { Checkbox, Icon, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';

import { IScopeNode } from '../models/IScopeNode';
import { MAX_FOLDERS_SHOWN, ScopeService } from '../services/ScopeService';
import styles from './ReInherit.module.scss';

const KIND_ICONS: { [kind: string]: string } = {
  web: 'Globe',
  library: 'DocLibrary',
  list: 'BulletedList',
  folder: 'FabricFolder'
};

interface IScopePickerProps {
  service: ScopeService;
  selected: Map<string, IScopeNode>;
  onToggle: (node: IScopeNode, checked: boolean) => void;
  disabled?: boolean;
}

function canExpand(node: IScopeNode): boolean {
  return node.kind === 'web' || node.kind === 'library' || (node.kind === 'folder' && (node.itemCount || 0) > 0);
}

function describe(node: IScopeNode): string | undefined {
  if (node.kind === 'library' || node.kind === 'list') {
    return `${(node.itemCount || 0).toLocaleString()} items`;
  }
  return undefined;
}

/**
 * Site collection -> subsites -> lists/libraries -> folders, loaded one level
 * at a time as nodes are expanded, so opening the picker on a large site
 * costs a couple of requests rather than a crawl. Lists are not expanded
 * into folders: list folders are rare, and a whole list can be selected.
 */
export const ScopePicker: React.FC<IScopePickerProps> = ({ service, selected, onToggle, disabled }) => {
  const [root, setRoot] = useState<IScopeNode | undefined>(undefined);
  const [error, setError] = useState<string | undefined>(undefined);
  const [expanded, setExpanded] = useState<Set<string>>(new Set<string>());
  // Bumped whenever children arrive so the tree re-renders; nodes are mutated in place.
  const [, setVersion] = useState(0);

  const loadChildren = useCallback(
    async (node: IScopeNode): Promise<void> => {
      node.loading = true;
      node.error = undefined;
      setVersion((v) => v + 1);
      try {
        node.children = await service.getChildren(node);
      } catch (err) {
        node.error = err instanceof Error ? err.message : 'Could not load';
      } finally {
        node.loading = false;
        setVersion((v) => v + 1);
      }
    },
    [service]
  );

  useEffect(() => {
    let active = true;
    service
      .getRootNode()
      .then(async (node) => {
        if (!active) {
          return;
        }
        setRoot(node);
        const initial = new Set<string>();
        initial.add(node.key);
        setExpanded(initial);
        await loadChildren(node);
      })
      .catch((err) => active && setError(err instanceof Error ? err.message : 'The site could not be loaded.'));
    return () => {
      active = false;
    };
  }, [service, loadChildren]);

  const toggleExpand = (node: IScopeNode): void => {
    const next = new Set<string>();
    expanded.forEach((key) => next.add(key));
    if (next.has(node.key)) {
      next.delete(node.key);
    } else {
      next.add(node.key);
      if (!node.children && !node.loading) {
        loadChildren(node).catch(() => undefined);
      }
    }
    setExpanded(next);
  };

  const renderNode = (node: IScopeNode, depth: number): React.ReactNode => {
    const isOpen = expanded.has(node.key);
    const expandable = canExpand(node);
    const detail = describe(node);
    return (
      <div key={node.key} role="treeitem" aria-expanded={expandable ? isOpen : undefined} aria-selected={selected.has(node.key)}>
        <div className={styles.pickerRow} style={{ paddingLeft: 6 + depth * 20 }}>
          {expandable ? (
            <button
              type="button"
              className={styles.chevronButton}
              onClick={() => toggleExpand(node)}
              aria-label={isOpen ? `Collapse ${node.title}` : `Expand ${node.title}`}
            >
              <Icon iconName={isOpen ? 'ChevronDown' : 'ChevronRight'} />
            </button>
          ) : (
            <span className={styles.chevronSpacer} />
          )}
          <Checkbox
            checked={selected.has(node.key)}
            disabled={disabled}
            onChange={(_, checked) => onToggle(node, !!checked)}
            ariaLabel={`Select ${node.title}`}
          />
          <span className={node.kind === 'web' ? styles.webBadge : styles.kindIcon}>
            <Icon iconName={KIND_ICONS[node.kind]} />
          </span>
          <span
            className={node.kind === 'web' ? styles.pickerTitleStrong : styles.pickerTitle}
            onClick={() => !disabled && onToggle(node, !selected.has(node.key))}
            title={node.serverRelativeUrl}
          >
            {node.title}
          </span>
          {node.isRootWeb && <span className={styles.pillNeutral}>Site collection root</span>}
          {node.hasUniquePermissions && !node.isRootWeb && <span className={styles.pillUnique}>Unique</span>}
          <span className={styles.spacer} />
          {detail && <span className={styles.pickerMeta}>{detail}</span>}
        </div>
        {isOpen && (
          <div role="group">
            {node.loading && (
              <div className={styles.pickerNote} style={{ paddingLeft: 42 + depth * 20 }}>
                <Spinner size={SpinnerSize.xSmall} label="Loading…" labelPosition="right" />
              </div>
            )}
            {node.error && (
              <div className={`${styles.pickerNote} ${styles.errorText}`} style={{ paddingLeft: 42 + depth * 20 }}>
                {node.error}
              </div>
            )}
            {node.children && node.children.length === 0 && !node.loading && (
              <div className={styles.pickerNote} style={{ paddingLeft: 42 + depth * 20 }}>
                {node.kind === 'web' ? 'No subsites or lists' : 'No subfolders'}
              </div>
            )}
            {node.children && node.children.map((child) => renderNode(child, depth + 1))}
            {node.kind === 'folder' || node.kind === 'library'
              ? node.children &&
                node.children.length >= MAX_FOLDERS_SHOWN && (
                  <div className={styles.pickerNote} style={{ paddingLeft: 42 + depth * 20 }}>
                    Showing the first {MAX_FOLDERS_SHOWN.toLocaleString()} folders. Select the parent to cover them all.
                  </div>
                )
              : undefined}
          </div>
        )}
      </div>
    );
  };

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }
  if (!root) {
    return <Spinner size={SpinnerSize.medium} label="Loading sites and libraries…" className={styles.pickerLoading} />;
  }
  return (
    <div className={styles.picker} role="tree" aria-label="Sites, lists and folders">
      {renderNode(root, 0)}
    </div>
  );
};
