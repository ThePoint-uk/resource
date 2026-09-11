import * as React from 'react';
import { Icon, DefaultButton } from '@fluentui/react';
import styles from './Shared.module.scss';

export const LoadingState: React.FC<{ rows?: number }> = ({ rows = 4 }) => (
  <div aria-busy="true" aria-label="Loading">
    {Array.from({ length: rows }).map((_, i) => (
      <div className={styles.shimmerRow} key={i}>
        <div className={styles.shimmerCircle} />
        <div className={styles.shimmerLines}>
          <div className={styles.shimmerLine} />
          <div className={styles.shimmerLine} />
        </div>
      </div>
    ))}
  </div>
);

export const ErrorState: React.FC<{ message: string; onRetry?: () => void; hint?: string }> = ({ message, onRetry, hint }) => (
  <div className={`${styles.state} ${styles.errorState}`} role="alert">
    <Icon iconName="Warning" className={styles.icon} />
    <div className={styles.title}>Couldn&apos;t load this widget</div>
    <div className={styles.message}>{message}</div>
    {hint && <div className={styles.message}>{hint}</div>}
    {onRetry && <DefaultButton text="Try again" iconProps={{ iconName: 'Refresh' }} onClick={onRetry} />}
  </div>
);

export const EmptyState: React.FC<{ icon?: string; title: string; message?: string; children?: React.ReactNode }> = ({ icon = 'Inbox', title, message, children }) => (
  <div className={styles.state}>
    <Icon iconName={icon} className={styles.icon} />
    <div className={styles.title}>{title}</div>
    {message && <div className={styles.message}>{message}</div>}
    {children}
  </div>
);

export const SectionHeader: React.FC<{ text: string }> = ({ text }) => <div className={styles.sectionHeader}>{text}</div>;

export const FooterLink: React.FC<{ href: string; text: string; newTab?: boolean }> = ({ href, text, newTab = true }) => (
  <div className={styles.footerLink}>
    <a href={href} target={newTab ? '_blank' : undefined} rel="noreferrer">
      {text} <Icon iconName="ChevronRight" styles={{ root: { fontSize: 10 } }} />
    </a>
  </div>
);
