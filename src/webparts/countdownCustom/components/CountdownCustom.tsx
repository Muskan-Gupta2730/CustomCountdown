import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './CountdownCustom.module.scss';
import { ICountdownCustomProps } from './ICountdownCustomProps';

interface IListRow {
  Title: string;
  DataTime?: string;
  DateTime?: string;
  ApplyLink: { Url: string; Description: string } | undefined;
  isActive: string;
  PageLinks: string | undefined;
}

interface ICountdownItem {
  title: string;
  link: string;
  targetMs: number;
}

const HIDE_AFTER_MS = 24 * 60 * 60 * 1000; // closed section stays for 24 hours

const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
const COLORS = {
  dark: '#150056',
  accent: '#b44ef0',
  panel: '#ffffff',
  card: '#ffffff',
};

const normalizePath = (value: string): string => {
  let p = (value || '').trim();
  if (!p) return '';
  try {
    p = new URL(p, window.location.origin).pathname;
  } catch {
    /* keep raw value */
  }
  try {
    p = decodeURIComponent(p);
  } catch {
    /* keep encoded value */
  }
  return p.replace(/\/+$/, '').toLowerCase();
};

const CountdownCustom: React.FC<ICountdownCustomProps> = (props) => {
  const [items, setItems] = React.useState<ICountdownItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [now, setNow] = React.useState<number>(Date.now());
  const [offset, setOffset] = React.useState<number>(0);

  React.useEffect(() => {
    if (!props.listTitle) {
      setItems([]);
      setError('');
      return;
    }

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
        if (!res.ok) {
          throw new Error(`List "${props.listTitle}" could not be read (${res.status}).`);
        }
        return res.json();
      })
      .then((json: { value: IListRow[] }) => {
        const currentPath = normalizePath(window.location.pathname);
        const isWorkbench = currentPath.indexOf('workbench.aspx') > -1;

        const matched = (json.value || []).filter((r) => {
          const pages = (r.PageLinks || '')
            .split(/[\s,;|]+/)
            .map(normalizePath)
            .filter((x) => x !== '');true; // empty = show on every page
          return isWorkbench || pages.indexOf(currentPath) > -1;
        });

        const list: ICountdownItem[] = matched
          .map((r) => ({
            title: r.Title,
            link: r.ApplyLink ? r.ApplyLink.Url : '',
                       targetMs: new Date(r.DateTime || r.DataTime || '').getTime(),
          }))
          .filter((i) => !isNaN(i.targetMs))
          .sort((a, b) => a.targetMs - b.targetMs);

        setItems(list);
        setLoading(false);
      })
      .catch((e: Error) => {
        setError(e.message);
        setLoading(false);
      });
  }, [props.listTitle, props.siteUrl]);

  React.useEffect(() => {
    const id = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(id);
  }, []);

  const rootStyle = {
    '--dark': COLORS.dark,
    '--accent': COLORS.accent,
    '--panel': COLORS.panel,
    '--card': COLORS.card,
  } as React.CSSProperties;

  const renderMessage = (text: string): JSX.Element => (
    <div className={styles.countdown} style={rootStyle}>
      <div className={styles.panel}>
        <div className={styles.content}>
          <div className={styles.title}>{text}</div>
        </div>
      </div>
    </div>
  );

  if (!props.listTitle) return renderMessage('Open the web part settings and enter the SharePoint list name.');
  if (loading && items.length === 0) return renderMessage('Loading...');
  if (error) return renderMessage(error);

  const serverNow = now + offset;

  // nearest upcoming -> else ended less than 24h ago -> else nothing (section vanishes)
  const upcoming = items.filter((i) => i.targetMs > serverNow)[0];
  const recentlyEnded = items
    .filter((i) => i.targetMs <= serverNow && serverNow - i.targetMs < HIDE_AFTER_MS)
    .sort((a, b) => b.targetMs - a.targetMs)[0];
  const item = upcoming || recentlyEnded;

  if (!item) {
    return props.isEditMode
      ? renderMessage('No countdown is scheduled for this page (visible only while editing).')
      : <></>;
  }

  const diff = Math.max(0, item.targetMs - serverNow);
  const expired = diff === 0;

  const units = [
    { v: Math.floor(diff / 86400000), l: 'DAYS' },
    { v: Math.floor((diff % 86400000) / 3600000), l: 'HRS' },
    { v: Math.floor((diff % 3600000) / 60000), l: 'MIN' },
    { v: Math.floor((diff % 60000) / 1000), l: 'SEC' },
  ];

  return (
    <div className={styles.countdown} style={rootStyle}>
      <div className={styles.panel}>
        {props.backgroundImageUrl && (
          <>
            <div
              className={styles.bgImage}
              style={{ backgroundImage: `url("${props.backgroundImageUrl}")` }}
            />
            <div
              className={styles.overlay}
              style={{ opacity: (props.overlayOpacity ?? 85) / 100 }}
            />
          </>
        )}

        <div className={styles.content}>
          <div className={styles.textBlock}>
            <div className={styles.label}>{props.label}</div>
            <div className={styles.title}>{item.title}</div>
          </div>

          {expired ? (
            <div className={styles.expired}>{props.expiredMessage}</div>
          ) : (
            <div className={styles.cards}>
              {units.map((u) => (
                <div className={styles.card} key={u.l}>
                  <div className={styles.num}>{pad(u.v)}</div>
                  <div className={styles.unit}>{u.l}</div>
                </div>
              ))}
            </div>
          )}

          {!expired && item.link && (
            <a
              className={styles.button}
              href={item.link}
              target="_blank"
              rel="noopener noreferrer"
              data-interception="off"
            >
              {props.buttonText}
            </a>
          )}
        </div>
      </div>
    </div>
  );
};

export default CountdownCustom;