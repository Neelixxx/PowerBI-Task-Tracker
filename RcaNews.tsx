import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IRcaNewsProps } from './IRcaNewsProps';
import { INewsItem } from './INewsItem';
import styles from './RcaNews.module.scss';

interface IState { items: INewsItem[]; loading: boolean; error?: string; }

export default class RcaNews extends React.Component<IRcaNewsProps, IState> {
  public state: IState = { items: [], loading: true };
  public componentDidMount(): void { this.loadNews().catch(() => undefined); }

  private async loadNews(): Promise<void> {
    const { context, maxItems, categoryFilter } = this.props;
    const webUrl = context.pageContext.web.absoluteUrl;
    const select = 'Id,Title,Description,FileRef,BannerImageUrl,FirstPublishedDate,Created,PromotedState,Author/Title,Category';
    let filter = 'PromotedState eq 2';
    if (categoryFilter.trim()) filter += ` and Category eq '${categoryFilter.trim().replace(/'/g, "''")}'`;
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('Site Pages')/items?$select=${encodeURIComponent(select)}&$expand=Author&$filter=${encodeURIComponent(filter)}&$orderby=FirstPublishedDate desc,Created desc&$top=${Math.max(4, Math.min(maxItems, 12))}`;
    try {
      const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!response.ok) throw new Error(`SharePoint returned ${response.status} ${response.statusText}`);
      const data = await response.json();
      const items: INewsItem[] = (data.value || []).map((row: any) => {
        let imageUrl: string | undefined;
        if (row.BannerImageUrl) imageUrl = typeof row.BannerImageUrl === 'string' ? row.BannerImageUrl : (row.BannerImageUrl.Url || row.BannerImageUrl.url);
        const rawUrl = row.FileRef || '';
        return {
          id: row.Id, title: row.Title || 'Untitled news post',
          description: row.Description || 'Read the latest RCA update.',
          url: rawUrl.startsWith('http') ? rawUrl : `${window.location.origin}${rawUrl}`,
          imageUrl, published: row.FirstPublishedDate || row.Created,
          category: row.Category || 'RCA Update', author: row.Author?.Title
        };
      });
      this.setState({ items, loading: false });
    } catch (e) {
      this.setState({ loading: false, error: e instanceof Error ? e.message : 'Unable to load SharePoint News.' });
    }
  }

  private date(value: string): string {
    return value ? new Intl.DateTimeFormat(undefined, { month:'short', day:'numeric', year:'numeric' }).format(new Date(value)) : '';
  }
  private image(item: INewsItem, hero=false): React.ReactNode {
    if (item.imageUrl) return <img className={hero ? styles.heroImage : styles.cardImage} src={item.imageUrl} alt="" />;
    return <div className={hero ? styles.heroFallback : styles.cardFallback} aria-hidden="true"><span>RCA</span></div>;
  }

  public render(): React.ReactElement<IRcaNewsProps> {
    const { items, loading, error } = this.state;
    if (loading) return <section className={styles.rcaNews}><div className={styles.status}>Loading RCA News…</div></section>;
    if (error) return <section className={styles.rcaNews}><div className={styles.error}><strong>RCA News could not load.</strong><span>{error}</span></div></section>;
    if (!items.length) return <section className={styles.rcaNews}><Header title={this.props.title} intro={this.props.intro}/><div className={styles.status}>No published SharePoint News posts were found.</div></section>;

    const featured=items[0], cards=items.slice(1);
    return <section className={styles.rcaNews}>
      <Header title={this.props.title} intro={this.props.intro}/>
      <a className={styles.hero} href={featured.url}>
        <div className={styles.heroMedia}>{this.image(featured,true)}</div>
        <div className={styles.heroContent}>
          <span className={styles.category}>{featured.category}</span>
          <h3>{featured.title}</h3><p>{featured.description}</p>
          <div className={styles.meta}><span>{this.date(featured.published)}</span>{featured.author && <span>• {featured.author}</span>}</div>
          <span className={styles.readMore}>Read featured story →</span>
        </div>
      </a>
      <div className={styles.grid}>{cards.map(item =>
        <a className={styles.card} href={item.url} key={item.id}>
          <div className={styles.cardMedia}>{this.image(item)}</div>
          <div className={styles.cardBody}>
            <span className={styles.category}>{item.category}</span>
            <h3>{item.title}</h3><p>{item.description}</p>
            <div className={styles.cardFooter}><span>{this.date(item.published)}</span><span>→</span></div>
          </div>
        </a>
      )}</div>
    </section>;
  }
}

const Header: React.FC<{title:string; intro:string}> = ({title,intro}) =>
  <header className={styles.sectionHeader}><div><div className={styles.eyebrow}>ROOT CAUSE ANALYSIS</div><h2>{title}</h2><p>{intro}</p></div><div className={styles.goldRule}/></header>;
