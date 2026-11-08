import * as React from 'react';
import { IListItem, IPerson, ISearchDocument } from '../models/Models';
import { formatDate, safeUrl, statusClass } from '../utils/helpers';

export const Stat: React.FC<{ label: string; value: number; note: string; icon: string }> = p => (
  <div className="es-stat">
    <div className="es-stat-top"><span>{p.label}</span><b aria-hidden="true">{p.icon}</b></div>
    <strong>{p.value}</strong>
    <small>{p.note}</small>
  </div>
);

export const PanelTitle: React.FC<{ title: string; subtitle: string }> = p => (
  <div className="es-panel-title"><div><h2>{p.title}</h2><p>{p.subtitle}</p></div></div>
);

export const ModuleHeader: React.FC<{ title: string; subtitle: string; icon: string; children?: React.ReactNode }> = p => (
  <div className="es-module-head">
    <div className="es-module-icon" aria-hidden="true">{p.icon}</div>
    <div><div className="es-eyebrow">EXPERIENCE SUITE</div><h1>{p.title}</h1><p>{p.subtitle}</p></div>
    {p.children}
  </div>
);

export const Empty: React.FC<{ text: string }> = p => (
  <div className="es-empty"><span aria-hidden="true">⌕</span><p>{p.text}</p></div>
);

export const ItemRow: React.FC<{ item: IListItem }> = ({ item }) => {
  const due = formatDate(item.DueDate);
  return (
    <div className="es-item-row">
      <div className="es-item-symbol" aria-hidden="true">▤</div>
      <div className="es-item-main">
        <strong>{item.Title}</strong>
        <small>{item.Description || item.Category || item.Owner || 'SharePoint list item'}</small>
      </div>
      <div className="es-item-meta">
        {item.Status && <span className={`es-status ${statusClass(item.Status)}`}>{item.Status}</span>}
        <small>{due ? `Due ${due}` : formatDate(item.Modified)}</small>
      </div>
    </div>
  );
};

export const SearchBar: React.FC<{ query: string; setQuery: (v: string) => void; onSearch: () => void; busy: boolean; placeholder: string; buttonText?: string }> = p => (
  <div className="es-searchbox">
    <span aria-hidden="true">⌕</span>
    <input
      value={p.query}
      aria-label={p.placeholder}
      onChange={e => p.setQuery(e.target.value)}
      onKeyDown={e => { if (e.key === 'Enter' && !p.busy) p.onSearch(); }}
      placeholder={p.placeholder}
    />
    <button onClick={p.onSearch} disabled={p.busy}>{p.busy ? 'Searching…' : (p.buttonText || 'Search')}</button>
  </div>
);

export const ResultsPanel: React.FC<{ searched: boolean; empty: string; items: ISearchDocument[] }> = p => (
  <section className="es-panel es-results">
    <PanelTitle title="Search results" subtitle={p.searched ? `${p.items.length} results` : 'No search run yet'} />
    {p.items.length
      ? p.items.map((x, i) => (
        <a className="es-result" href={safeUrl(x.path)} target="_blank" rel="noopener noreferrer" key={`${x.path}-${i}`}>
          <div className="es-result-icon" aria-hidden="true">{x.fileType === 'Site' ? '◫' : '▤'}</div>
          <div>
            <strong>{x.title}</strong>
            <small>{x.path}</small>
            <small>{x.author}{formatDate(x.modified) ? ` · ${formatDate(x.modified)}` : ''}</small>
          </div>
          <b aria-hidden="true">↗</b>
        </a>
      ))
      : <Empty text={p.searched ? 'No results matched your search.' : p.empty} />}
  </section>
);

export const PersonCard: React.FC<{ person: IPerson }> = ({ person: p }) => {
  const picture = safeUrl(p.picture);
  const profile = safeUrl(p.profileUrl);
  return (
    <article className="es-person-card">
      <div className="es-person-top">
        {picture ? <img src={picture} alt="" /> : <div className="es-person-avatar">{(p.title.charAt(0) || '?').toUpperCase()}</div>}
        <div><strong>{p.title}</strong><small>{p.jobTitle || 'Colleague'}</small></div>
      </div>
      <div className="es-person-details">
        {p.department && <span>▦ {p.department}</span>}
        {p.location && <span>⌖ {p.location}</span>}
        {p.email && <a href={`mailto:${p.email}`}>✉ {p.email}</a>}
      </div>
      {p.skills && (
        <div className="es-skills">
          {p.skills.split(/[;,|]/).map(s => s.trim()).filter(Boolean).slice(0, 5).map((s, i) => <span key={i}>{s}</span>)}
        </div>
      )}
      {profile && <a className="es-profile-link" href={profile} target="_blank" rel="noopener noreferrer">Open profile ↗</a>}
    </article>
  );
};
