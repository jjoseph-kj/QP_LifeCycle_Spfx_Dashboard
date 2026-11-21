import * as React from 'react';
import styles from './Qplifecycle.module.scss';
import type { IQpcycleProps } from './IQplifecycleProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IQpItem {
  id: number;
  qpId: string;
  status: string;
  changeDate: Date | undefined;
  days: number | undefined;
}

type SortKey = 'qpId' | 'status' | 'changeDate' | 'days';

interface IState {
  items: IQpItem[];
  loading: boolean;
  error: string;
  search: string;
  statusFilter: string;
  sortKey: SortKey;
  sortAsc: boolean;
  lastRefreshed: Date | undefined;
}

const STATUSES: string[] = ['Draft QP', 'Proposed QP', 'Agreed QP'];
const STATUS_CLASS: { [k: string]: string } = {
  'Draft QP': 'draft',
  'Proposed QP': 'proposed',
  'Agreed QP': 'agreed'
};
const MS_PER_DAY: number = 1000 * 60 * 60 * 24;

const fmtDate = (d: Date | undefined): string =>
  d ? d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

// Calendar-day number (DST-safe): uses Y/M/D only, so 23h/25h days can't skew the result
const startOfDay = (d: Date): number => Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());

// Date-only values arrive as "YYYY-MM-DDT00:00:00Z": take Y/M/D literally (no timezone shift).
// Values with a real time part are converted to local time as normal.
const parseSpDate = (raw: string | undefined | null): Date | undefined => {
  if (!raw) { return undefined; }
  const m: RegExpMatchArray | null = /^(\d{4})-(\d{2})-(\d{2})T00:00:00(?:\.0+)?Z$/.exec(raw);
  if (m) { return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); }
  const d: Date = new Date(raw);
  return isNaN(d.getTime()) ? undefined : d;
};
export default class Qpcycle extends React.Component<IQpcycleProps, IState> {
  constructor(props: IQpcycleProps) {
    super(props);
    this.state = {
      items: [],
      loading: true,
      error: '',
      search: '',
      statusFilter: 'All',
      sortKey: 'days',
      sortAsc: false,
      lastRefreshed: undefined
    };
  }

   public componentDidMount(): void {
    this._load().catch(() => undefined);
  }

  private _load = async (): Promise<void> => {
    this.setState({ loading: true, error: '' });
    try {
      const { siteUrl, listName, spHttpClient } = this.props;
      const select: string = 'Id,QP_x0020_ID,Lifecycle_x0020_Change_x0020_Dat,field_2';
      let url: string =
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items` +
        `?$select=${select}&$top=5000`;
      const all: IQpItem[] = [];
      const today: number = startOfDay(new Date());

      while (url) {
        const res: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!res.ok) {
          throw new Error(`${res.status} ${res.statusText}`);
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const json: any = await res.json();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (json.value as any[]).forEach((r: any) => {
          const raw: string | undefined = r.Lifecycle_x0020_Change_x0020_Dat;
          const d: Date | undefined = parseSpDate(raw);
          all.push({
            id: r.Id,
            qpId: r.QP_x0020_ID ? String(r.QP_x0020_ID) : '',
            status: r.field_2 ? String(r.field_2) : '',
            changeDate: d,
            days: d ? Math.round((today - startOfDay(d)) / MS_PER_DAY) : undefined
          });
        });
        url = json['odata.nextLink'] || json['@odata.nextLink'] || '';
      }

      this.setState({ items: all, loading: false, lastRefreshed: new Date() });
    } catch (e) {
      this.setState({ loading: false, error: (e as Error).message });
    }
  };

  private _visibleItems(): IQpItem[] {
    const { items, search, statusFilter, sortKey, sortAsc } = this.state;
    const q: string = search.trim().toLowerCase();
    const list: IQpItem[] = items.filter(i =>
      (statusFilter === 'All' || i.status === statusFilter) &&
      (!q || i.qpId.toLowerCase().indexOf(q) > -1)
    );
    const dir: number = sortAsc ? 1 : -1;
    list.sort((a, b) => {
      let av: string | number;
      let bv: string | number;
      switch (sortKey) {
        case 'qpId': av = a.qpId.toLowerCase(); bv = b.qpId.toLowerCase(); break;
        case 'status': av = a.status; bv = b.status; break;
        case 'changeDate': av = a.changeDate ? a.changeDate.getTime() : -1; bv = b.changeDate ? b.changeDate.getTime() : -1; break;
        default: av = a.days === undefined ? -1 : a.days; bv = b.days === undefined ? -1 : b.days;
      }
      return av < bv ? -1 * dir : av > bv ? 1 * dir : 0;
    });
    return list;
  }

  private _sort = (key: SortKey): void => {
    this.setState(s => ({ sortKey: key, sortAsc: s.sortKey === key ? !s.sortAsc : true }));
  };

  private _export = (): void => {
    const rows: IQpItem[] = this._visibleItems();
    const esc = (v: string): string =>
      v.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const head: string = ['QP ID', 'Status', 'Lifecycle Change Date', 'No. of Days']
      .map(h => `<th style="background:#0f4c81;color:#fff;font-weight:bold;border:1px solid #999">${h}</th>`).join('');
    const body: string = rows.map(r =>
      `<tr><td style="border:1px solid #ccc">${esc(r.qpId)}</td>` +
      `<td style="border:1px solid #ccc">${esc(r.status)}</td>` +
      `<td style="border:1px solid #ccc">${esc(fmtDate(r.changeDate))}</td>` +
      `<td style="border:1px solid #ccc">${r.days === undefined ? '' : r.days}</td></tr>`
    ).join('');
    const html: string =
      '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<head><meta charset="utf-8"></head>' +
      `<body><table>${`<tr>${head}</tr>`}${body}</table></body></html>`;
    const blob: Blob = new Blob(['﻿' + html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    const stamp: string = new Date().toISOString().slice(0, 10);
    const a: HTMLAnchorElement = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `LifeCycle_Tracker_${stamp}.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  };

  private _badge(status: string): React.ReactElement {
    const key: string = STATUS_CLASS[status] || 'other';
    return <span className={`${styles.badge} ${styles[key]}`}>{status || 'Unassigned'}</span>;
  }

  private _daysPill(days: number | undefined): React.ReactElement {
    if (days === undefined) { return <span>—</span>; }
    const cls: string = days > 90 ? styles.dRed : days > 30 ? styles.dAmber : styles.dGreen;
    return <span className={`${styles.daysPill} ${cls}`}>{days} {days === 1 ? 'day' : 'days'}</span>;
  }

  private _th(label: string, key: SortKey): React.ReactElement {
    const { sortKey, sortAsc } = this.state;
    return (
      <th onClick={() => this._sort(key)} className={styles.sortable}>
        {label} <span className={styles.arrow}>{sortKey === key ? (sortAsc ? '▲' : '▼') : '↕'}</span>
      </th>
    );
  }

  public render(): React.ReactElement<IQpcycleProps> {
    const { items, loading, error, search, statusFilter, lastRefreshed } = this.state;
    const rows: IQpItem[] = this._visibleItems();
    const count = (s: string): number => items.filter(i => i.status === s).length;

    return (
      <section className={styles.qpcycle}>
        <div className={styles.header}>
          <div>
            <h2 className={styles.title}>LifeCycle Tracker</h2>
            <div className={styles.subtitle}>
              Quality Plans lifecycle status and ageing
              {lastRefreshed ? ` · Last refreshed ${lastRefreshed.toLocaleTimeString()}` : ''}
            </div>
          </div>
          <div className={styles.actions}>
            <button className={styles.btn} onClick={() => { this._load().catch(() => undefined); }} disabled={loading}>
              ⟳ Refresh
            </button>
            <button className={`${styles.btn} ${styles.btnPrimary}`} onClick={this._export} disabled={loading || rows.length === 0}>
              ⬇ Export to Excel
            </button>
          </div>
        </div>

        <div className={styles.cards}>
          <div className={`${styles.card} ${styles.cTotal}`} onClick={() => this.setState({ statusFilter: 'All' })}>
            <div className={styles.cardNum}>{items.length}</div><div className={styles.cardLbl}>Total QPs</div>
          </div>
          {STATUSES.map(s => (
            <div key={s} className={`${styles.card} ${styles['c_' + STATUS_CLASS[s]]}`} onClick={() => this.setState({ statusFilter: s })}>
              <div className={styles.cardNum}>{count(s)}</div><div className={styles.cardLbl}>{s}</div>
            </div>
          ))}
        </div>

        <div className={styles.toolbar}>
          <input
            className={styles.search}
            placeholder="Search QP ID..."
            value={search}
            onChange={e => this.setState({ search: e.target.value })}
          />
          <select className={styles.select} value={statusFilter} onChange={e => this.setState({ statusFilter: e.target.value })}>
            <option value="All">All statuses</option>
            {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <span className={styles.count}>{rows.length} of {items.length} records</span>
        </div>

        {error && <div className={styles.error}>Could not load data: {error}</div>}

        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                {this._th('QP ID', 'qpId')}
                {this._th('Status', 'status')}
                {this._th('Lifecycle Change Date', 'changeDate')}
                {this._th('No. of Days', 'days')}
              </tr>
            </thead>
            <tbody>
              {loading && <tr><td colSpan={4} className={styles.empty}>Loading…</td></tr>}
              {!loading && rows.length === 0 && <tr><td colSpan={4} className={styles.empty}>No records found</td></tr>}
              {!loading && rows.map(r => (
                <tr key={r.id}>
                  <td className={styles.qpId}>{r.qpId}</td>
                  <td>{this._badge(r.status)}</td>
                  <td>{fmtDate(r.changeDate)}</td>
                  <td>{this._daysPill(r.days)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className={styles.legend}>
          Days in current stage: <span className={`${styles.daysPill} ${styles.dGreen}`}>≤ 30</span>
          <span className={`${styles.daysPill} ${styles.dAmber}`}>31–90</span>
          <span className={`${styles.daysPill} ${styles.dRed}`}>&gt; 90</span>
        </div>
      </section>
    );
  }
}
