import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './CountdownCustom.module.scss';
import { ICountdownCustomProps } from './ICountdownCustomProps';

// ─── SharePoint list row shape ────────────────────────────────────────────────
interface IListRow {
  Title: string;
  DataTime?: string;
  DateTime?: string;
  Description?: string;
  ApplyLink?: { Url: string; Description: string };
  isActive: string;
  PageLinks?: string;
}

interface ICountdownItem {
  title: string;
  description: string;
  link: string;
  targetMs: number;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
const HIDE_AFTER_MS = 24 * 60 * 60 * 1000; // keep "closed" card for 24 h

const pad = (n: number): string => (n < 10 ? '0' + n : '' + n);

/** Format a Date as "D MMM YYYY" – e.g. "31 OCT 2026" */
const formatCloses = (ms: number): string => {
  const d = new Date(ms);
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
};

const normalizePath = (value: string): string => {
  let p = (value || '').trim();
  if (!p) return '';
  try { p = new URL(p, window.location.origin).pathname; } catch { /* keep */ }
  try { p = decodeURIComponent(p); } catch { /* keep */ }
  return p.replace(/\/+$/, '').toLowerCase();
};

// ─────────────────────────────────────────────────────────────────────────────
const CountdownCustom: React.FC<ICountdownCustomProps> = (props) => {
  const [items, setItems]   = React.useState<ICountdownItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError]   = React.useState<string>('');
  const [now, setNow]       = React.useState<number>(Date.now());
  const [offset, setOffset] = React.useState<number>(0);

  // ── Fetch list data ────────────────────────────────────────────────────────
  React.useEffect(() => {
    if (!props.listTitle) { setItems([]); setError(''); return; }

    setLoading(true);
    setError('');

    const safeTitle = encodeURIComponent(props.listTitle.replace(/'/g, "''"));
    const url =
      `${props.siteUrl}/_api/web/lists/getbytitle('${safeTitle}')/items` +
      `?$filter=isActive eq 'Yes'&$top=100`;

    props.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => {
        const serverDate = res.headers.get('Date');
        if (serverDate) {
          const s = new Date(serverDate).getTime();
          if (!isNaN(s)) setOffset(s - Date.now());
        }
        if (!res.ok) throw new Error(`List "${props.listTitle}" could not be read (${res.status}).`);
        return res.json();
      })
      .then((json: { value: IListRow[] }) => {
        const currentPath = normalizePath(window.location.pathname);
        const isWorkbench  = currentPath.indexOf('workbench.aspx') > -1;

        const matched = (json.value || []).filter((r) => {
          const pages = (r.PageLinks || '')
            .split(/[\s,;|]+/)
            .map(normalizePath)
            .filter((x) => x !== '');
          return isWorkbench || pages.length === 0 || pages.indexOf(currentPath) > -1;
        });

        const list: ICountdownItem[] = matched
          .map((r) => ({
            title:       r.Title,
            description: r.Description || '',
            link:        r.ApplyLink ? r.ApplyLink.Url : '',
            targetMs:    new Date(r.DateTime || r.DataTime || '').getTime(),
          }))
          .filter((i) => !isNaN(i.targetMs))
          .sort((a, b) => a.targetMs - b.targetMs);

        setItems(list);
        setLoading(false);
      })
      .catch((e: Error) => { setError(e.message); setLoading(false); });
  }, [props.listTitle, props.siteUrl]);

  // ── 1-second ticker ───────────────────────────────────────────────────────
  React.useEffect(() => {
    const id = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(id);
  }, []);

  // ── Simple message helper ─────────────────────────────────────────────────
  const renderMessage = (text: string): JSX.Element => (
    <div className={styles.wrapper}>
      <div className={styles.card}>
        <div className={styles.message}>{text}</div>
      </div>
    </div>
  );

  if (!props.listTitle)          return renderMessage('Open the web part settings and enter the SharePoint list name.');
  if (loading && !items.length)  return renderMessage('Loading…');
  if (error)                     return renderMessage(error);

  const serverNow = now + offset;

  const upcoming      = items.filter((i) => i.targetMs > serverNow)[0];
  const recentlyEnded = items
    .filter((i) => i.targetMs <= serverNow && serverNow - i.targetMs < HIDE_AFTER_MS)
    .sort((a, b) => b.targetMs - a.targetMs)[0];
  const item = upcoming || recentlyEnded;

  if (!item) {
    return props.isEditMode
      ? renderMessage('No active countdown for this page (only visible in edit mode).')
      : <></>;
  }

  const diff    = Math.max(0, item.targetMs - serverNow);
  const expired = diff === 0;

  const days  = Math.floor(diff / 86_400_000);
  const hours = Math.floor((diff % 86_400_000) / 3_600_000);
  const mins  = Math.floor((diff % 3_600_000)  / 60_000);
  const secs  = Math.floor((diff % 60_000)     / 1_000);

  const units = [
    { v: days,  l: 'DAYS'  },
    { v: hours, l: 'HOURS' },
    { v: mins,  l: 'MINS'  },
    { v: secs,  l: 'SECS'  },
  ];

  // Description: prefer list data, fallback to webpart property
  const descriptionText = item.description || props.description || '';

  return (
    <div className={styles.wrapper}>
      <div className={styles.card}>

        {/* ── Optional background image + overlay ─────────────────────────── */}
        {props.backgroundImageUrl && (
          <>
            <div
              className={styles.bgImage}
              style={{ backgroundImage: `url("${props.backgroundImageUrl}")` }}
            />
            <div
              className={styles.bgOverlay}
              style={{ opacity: (props.overlayOpacity ?? 85) / 100 }}
            />
          </>
        )}

        {/* ── Main content row ─────────────────────────────────────────────── */}
        <div className={styles.inner}>

          {/* LEFT – text block */}
          <div className={styles.textBlock}>
            <p className={styles.topLabel}>{props.label || 'Applications close in'}</p>
            <h2 className={styles.title}>{item.title}</h2>
            {descriptionText && (
              <p className={styles.description}>{descriptionText}</p>
            )}
            {!expired && item.link && (
              <a
                className={styles.applyBtn}
                href={item.link}
                target="_blank"
                rel="noopener noreferrer"
                data-interception="off"
              >
                {props.buttonText || 'Apply now'}
                <span className={styles.arrow} aria-hidden="true">›</span>
              </a>
            )}
          </div>

          {/* RIGHT – countdown box */}
          <div className={styles.countdownBox}>
            <span className={styles.closesLabel}>
              Closes {formatCloses(item.targetMs)}
            </span>

            {expired ? (
              <div className={styles.expiredText}>{props.expiredMessage || 'Applications are now closed'}</div>
            ) : (
              <div className={styles.tilesRow}>
                {units.map((u, idx) => (
                  <React.Fragment key={u.l}>
                    <div className={styles.tile}>
                      <span className={styles.tileNum}>{pad(u.v)}</span>
                      <span className={styles.tileLabel}>{u.l}</span>
                    </div>
                    {idx < units.length - 1 && (
                      <span className={styles.separator} aria-hidden="true"></span>
                    )}
                  </React.Fragment>
                ))}
              </div>
            )}
          </div>

        </div>
      </div>
    </div>
  );
};

export default CountdownCustom;
