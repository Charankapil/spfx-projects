import * as React from 'react';
import { Panel, PanelType, Icon } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';

interface ITopic {
  icon: string;
  title: string;
  steps: React.ReactNode[];
}

const TOPICS: ITopic[] = [
  {
    icon: 'Nav2DMapView',
    title: 'Getting around',
    steps: [
      <>The <b>left menu</b> has your personal pages at the top and every board below. Click <b>«</b> to make it narrow.</>,
      <><b>Home</b> shows what needs your attention. <b>My Work</b> lists everything assigned to you.</>,
      <><b>My projects</b> lists boards where you are project head, lead, sponsor or owner. <b>My team</b> shows your reports&apos; work.</>,
      <>Lost? The line at the top (for example <i>Home › Marketing › Q4 Campaign</i>) shows where you are. Click any part to go back.</>,
      <>Fastest of all: press <kbd className={styles.kbd}>Ctrl</kbd> + <kbd className={styles.kbd}>K</kbd> (or click <b>Search or jump to</b>) and type part of a board&apos;s name.</>
    ]
  },
  {
    icon: 'BacklogBoard',
    title: 'Boards',
    steps: [
      <>A <b>board</b> is one project or one team&apos;s list of work. Click <b>+ New board</b> and pick a template to start.</>,
      <>Items are grouped into coloured <b>groups</b>, such as <i>Planning</i> or <i>This week</i>. Rename a group by clicking its name.</>,
      <>Each column holds one kind of detail: who owns it, its status, its dates. Click <b>+</b> at the end of the header row to add a column.</>,
      <>Click the <b>☆ star</b> next to a board&apos;s name to keep it at the top of the left menu.</>
    ]
  },
  {
    icon: 'CheckboxComposite',
    title: 'Items (the rows)',
    steps: [
      <>Type in <b>+ Add item</b> at the bottom of a group and press Enter.</>,
      <>Click any cell to change it: pick a status colour, a person or a date.</>,
      <>Click the <b>speech bubble</b> on a row to open the item. There you can post <b>updates</b> (type @ to mention someone), add or link <b>files</b>, and see the <b>activity</b> history.</>,
      <>Drag a row to reorder it or move it to another group. Tick several rows to move or delete them together.</>
    ]
  },
  {
    icon: 'ViewAll',
    title: 'Views',
    steps: [
      <><b>Main table</b>: every item as rows and columns, like a spreadsheet.</>,
      <><b>Kanban</b>: cards in columns by status. Drag a card to change its status.</>,
      <><b>Timeline</b>: a calendar bar for each item. Drag a bar to move it, or drag its ends to change the dates.</>,
      <>Use <b>Search</b>, <b>Person</b> and <b>Filter</b> above any view to see only what you need. <b>Clear all</b> brings everything back.</>
    ]
  },
  {
    icon: 'Lock',
    title: 'Who can see what',
    steps: [
      <>A normal board is open to everyone on this site. A <b>private</b> board (🔒) is only open to the people added to it.</>,
      <>Board owners can change columns and settings. Open <b>Settings</b> on a board to change its name, owners, project roles or privacy.</>,
      <>Deleted items, files and boards go to the site <b>recycle bin</b>, so they can be restored.</>
    ]
  }
];

const SHORTCUTS: [string[], string][] = [
  [['Ctrl', 'K'], 'Search or jump to a board or page'],
  [['?'], 'Open this help'],
  [['Enter'], 'Save what you typed in a cell'],
  [['Esc'], 'Cancel typing, or close a panel'],
  [['Ctrl', 'Enter'], 'Post an update']
];

/** In-app help: short, plain-language steps for the main tasks. */
export function HelpPanel(props: { onClose: () => void; onTour?: () => void }): JSX.Element {
  const [open, setOpen] = React.useState<string>(TOPICS[0].title);
  return (
    <Panel isOpen type={PanelType.medium} onDismiss={props.onClose} headerText="How to use Work Boards" closeButtonAriaLabel="Close help" isLightDismiss>
      <p className={styles.muted} style={{ marginTop: 4 }}>Pick a topic. Each one is a few short steps.</p>
      <div className={styles.helpTopics}>
        {TOPICS.map(t => {
          const isOpen = open === t.title;
          return (
            <section key={t.title} className={styles.helpTopic}>
              <button type="button" className={styles.helpTopicHead} aria-expanded={isOpen} onClick={() => setOpen(isOpen ? '' : t.title)}>
                <span className={styles.helpIcon}><Icon iconName={t.icon} /></span>
                <span className={styles.navLabel}>{t.title}</span>
                <Icon iconName={isOpen ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 11 }} />
              </button>
              {isOpen && (
                <ol className={styles.helpSteps}>
                  {t.steps.map((s, i) => <li key={i}>{s}</li>)}
                </ol>
              )}
            </section>
          );
        })}
      </div>
      <h3 className={styles.sectionTitle} style={{ marginTop: 20 }}>Keyboard shortcuts</h3>
      <table className={styles.simpleTable} style={{ marginTop: 6 }}>
        <tbody>
          {SHORTCUTS.map(([keys, what]) => (
            <tr key={what}>
              <td style={{ whiteSpace: 'nowrap', width: 130 }}>{keys.map((k, i) => <React.Fragment key={k}>{i > 0 && ' + '}<kbd className={styles.kbd}>{k}</kbd></React.Fragment>)}</td>
              <td>{what}</td>
            </tr>
          ))}
        </tbody>
      </table>
      {props.onTour && (
        <p style={{ marginTop: 16 }}>
          <button type="button" className={styles.linkBtn} onClick={props.onTour}><Icon iconName="Lightbulb" /> Show the getting-started steps on Home again</button>
        </p>
      )}
    </Panel>
  );
}
