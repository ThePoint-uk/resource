import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './Shared.module.scss';

export interface IItemRowProps {
  title: string;
  meta?: React.ReactNode;
  trailing?: React.ReactNode;
  icon?: string;
  iconColor?: string;
  imageUrl?: string;
  href?: string;
  newTab?: boolean;
  onClick?: () => void;
  ariaLabel?: string;
  leading?: React.ReactNode;
}

/**
 * Generic list row: icon or image, title/meta and optional trailing content.
 * Renders as a link when `href` is provided, otherwise a button.
 */
export const ItemRow: React.FC<IItemRowProps> = (props) => {
  const { title, meta, trailing, icon, iconColor, imageUrl, href, newTab = true, onClick, ariaLabel, leading } = props;
  const content: React.ReactNode = (
    <>
      {leading}
      {!leading && imageUrl && <img className={styles.rowImage} src={imageUrl} alt="" loading="lazy" />}
      {!leading && !imageUrl && icon && (
        <span className={styles.rowIcon}>
          <Icon iconName={icon} styles={{ root: { color: iconColor } }} />
        </span>
      )}
      <span className={styles.rowBody}>
        <span className={styles.rowTitle} title={title}>{title}</span>
        {meta && <span className={styles.rowMeta}>{meta}</span>}
      </span>
      {trailing && <span className={styles.rowTrailing}>{trailing}</span>}
    </>
  );
  if (href) {
    return (
      <a className={styles.row} href={href} target={newTab ? '_blank' : undefined} rel="noreferrer" aria-label={ariaLabel || title} onClick={onClick}>
        {content}
      </a>
    );
  }
  return (
    <button type="button" className={styles.row} onClick={onClick} aria-label={ariaLabel || title}>
      {content}
    </button>
  );
};

export const ItemList: React.FC<{ compact?: boolean; children?: React.ReactNode }> = ({ compact, children }) => (
  <div className={`${styles.list} ${compact ? styles.compact : ''}`}>{children}</div>
);
