import * as React from 'react';
import { useState } from 'react';
import { Icon, Text } from '@fluentui/react';

import { ILibraryNode } from '../models/ILibraryNode';
import { IWebNode } from '../models/IWebNode';import styles from './FileTypeAnalyser.module.scss';

export interface IWebNodeTreeProps {
  node: IWebNode;
  depth: number;
}

export const WebNodeTree: React.FC<IWebNodeTreeProps> = ({ node, depth }) => {
  const [expanded, setExpanded] = useState(depth < 1);
  const hasChildren = node.webs.length > 0 || node.libraries.length > 0;

  return (
    <div className={styles.treeNode} style={{ marginLeft: depth * 16 }}>
      <div className={styles.treeNodeRow} onClick={() => setExpanded(!expanded)}>
        {hasChildren ? (
          <Icon iconName={expanded ? 'ChevronDown' : 'ChevronRight'} className={styles.treeChevron} />
        ) : (
          <span className={styles.treeChevronSpacer} />
        )}
        <Icon iconName="SharepointLogo" className={styles.treeIcon} />
        <Text variant="mediumPlus">{node.title || node.serverRelativeUrl}</Text>
        {node.error && (
          <Text variant="small" className={styles.errorLabel}>
            {node.error}
          </Text>
        )}
      </div>

      {expanded && (
        <div>
          {node.libraries.map((lib) => (
            <LibraryRow key={lib.id} library={lib} depth={depth + 1} />
          ))}
          {node.webs.map((child) => (
            <WebNodeTree key={child.id} node={child} depth={depth + 1} />
          ))}
        </div>
      )}
    </div>
  );
};

const LibraryRow: React.FC<{ library: ILibraryNode; depth: number }> = ({ library, depth }) => {
  const [expanded, setExpanded] = useState(false);
  const hasTypes = library.fileTypes.length > 0;

  return (
    <div className={styles.treeNode} style={{ marginLeft: depth * 16 }}>
      <div className={styles.treeNodeRow} onClick={() => hasTypes && setExpanded(!expanded)}>
        {hasTypes ? (
          <Icon iconName={expanded ? 'ChevronDown' : 'ChevronRight'} className={styles.treeChevron} />
        ) : (
          <span className={styles.treeChevronSpacer} />
        )}
        <Icon iconName="DocLibrary" className={styles.treeIcon} />
        <Text variant="medium">{library.title}</Text>
        {!library.scanned && (
          <Text variant="small" className={styles.pendingLabel}>
            &nbsp;(pending scan)
          </Text>
        )}
        {library.scanned && !library.error && (
          <Text variant="small" className={styles.countLabel}>
            {library.totalFiles.toLocaleString()} files            {library.totalFiles === 0 && library.itemCount > 0 &&
              ` (${library.itemCount.toLocaleString()} items not in search index)`}
          </Text>
        )}
        {library.error && (
          <Text variant="small" className={styles.errorLabel}>
            {library.error}
          </Text>
        )}
      </div>
      {expanded && hasTypes && (
        <div className={styles.fileTypeList} style={{ marginLeft: (depth + 1) * 16 }}>
          {library.fileTypes.map((stat) => (
            <div key={stat.extension} className={styles.fileTypeRow}>
              <Icon iconName="Page" className={styles.treeIcon} />
              <Text variant="small">.{stat.extension}</Text>
              <Text variant="small" className={styles.countLabel}>
                {stat.count.toLocaleString()}
              </Text>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
