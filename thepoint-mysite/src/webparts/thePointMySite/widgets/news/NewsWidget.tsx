import * as React from 'react';
import type { IWidgetDefinition, IWidgetProps } from '../../models';
import { useAsync } from '../../hooks/useAsync';
import { LoadingState, ErrorState, EmptyState } from '../../components/shared/States';
import { ItemList, ItemRow } from '../../components/shared/ItemRow';
import type { ISPNewsItem } from '../../services/SharePointService';
import { formatRelative } from '../../utils/format';
import styles from './NewsWidget.module.scss';

const NewsWidget: React.FC<IWidgetProps> = ({ sp, settings, density, size }) => {
  const max: number = Number(settings.maxItems) || 6;
  const scope: 'tenant' | 'site' | 'hub' = (String(settings.scope || 'tenant') as 'tenant' | 'site' | 'hub');
  const layout: string = String(settings.layout || 'auto');
  const news = useAsync<ISPNewsItem[]>(() => sp.getNews(max, scope), [sp, max, scope]);

  if (news.loading && !news.data) {
    return <LoadingState rows={4} />;
  }
  if (news.error) {
    return <ErrorState message={news.error} onRetry={news.reload} />;
  }
  if (!news.data || news.data.length === 0) {
    return <EmptyState icon="News" title="No news yet" message="Published news posts will appear here." />;
  }
  const useCards: boolean = layout === 'cards' || (layout === 'auto' && size >= 2);
  if (useCards) {
    return (
      <div className={styles.cards}>
        {news.data.map(n => (
          <a key={n.id} className={styles.card} href={n.url} target="_blank" rel="noreferrer">
            {n.imageUrl ? <img className={styles.image} src={n.imageUrl} alt="" loading="lazy" /> : <div className={styles.placeholder} />}
            <span className={styles.cardBody}>
              <span className={styles.cardTitle}>{n.title}</span>
              <span className={styles.cardMeta}>{[n.siteTitle, n.author, formatRelative(n.published)].filter(Boolean).join(' · ')}</span>
            </span>
          </a>
        ))}
      </div>
    );
  }
  return (
    <ItemList compact={density === 'compact'}>
      {news.data.map(n => (
        <ItemRow key={n.id} title={n.title} href={n.url} imageUrl={n.imageUrl} icon="News" meta={[n.siteTitle, n.author, formatRelative(n.published)].filter(Boolean).join(' · ')} />
      ))}
    </ItemList>
  );
};

export const newsDefinition: IWidgetDefinition = {
  type: 'news',
  title: 'News',
  description: 'Latest SharePoint news from across the organisation, this hub or this site.',
  icon: 'News',
  category: 'content',
  defaultSize: 2,
  allowMultiple: true,
  permissions: ['SharePoint search'],
  settingsSchema: [
    { key: 'maxItems', label: 'Number of posts', type: 'number', defaultValue: 6, min: 2, max: 20 },
    { key: 'scope', label: 'Scope', type: 'choice', defaultValue: 'tenant', choices: [{ key: 'tenant', text: 'Whole organisation' }, { key: 'hub', text: 'This hub' }, { key: 'site', text: 'This site' }] },
    { key: 'layout', label: 'Layout', type: 'choice', defaultValue: 'auto', choices: [{ key: 'auto', text: 'Automatic' }, { key: 'cards', text: 'Cards' }, { key: 'list', text: 'List' }] }
  ],
  component: NewsWidget
};

export default NewsWidget;
