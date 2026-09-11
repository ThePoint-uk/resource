import * as React from 'react';
import styles from './Shared.module.scss';
import { usePhoto } from '../../hooks/usePhoto';
import type { GraphService } from '../../services/GraphService';
import { initials } from '../../utils/format';

export interface IAvatarProps {
  graph: GraphService;
  userId?: string;
  name?: string;
  size?: 'sm' | 'md' | 'lg' | 'xl';
  presence?: string;
  photoUrl?: string;
}

const PALETTE: string[] = ['#0078d4', '#8764b8', '#038387', '#ca5010', '#b4009e', '#498205', '#c239b3', '#005b70', '#8e562e', '#4f6bed'];

function colorFor(name: string | undefined): string {
  if (!name) {
    return PALETTE[0];
  }
  let hash: number = 0;
  for (let i: number = 0; i < name.length; i++) {
    hash = (hash * 31 + name.charCodeAt(i)) | 0;
  }
  return PALETTE[Math.abs(hash) % PALETTE.length];
}

export function presenceClass(availability: string | undefined): string {
  switch ((availability || '').toLowerCase()) {
    case 'available':
    case 'availableidle':
      return styles.presenceAvailable;
    case 'busy':
    case 'busyidle':
    case 'inacall':
    case 'inameeting':
    case 'inaconferencecall':
      return styles.presenceBusy;
    case 'donotdisturb':
    case 'presenting':
      return styles.presenceDnd;
    case 'away':
    case 'berightback':
      return styles.presenceAway;
    default:
      return styles.presenceOffline;
  }
}

export const Avatar: React.FC<IAvatarProps> = ({ graph, userId, name, size = 'md', presence, photoUrl }) => {
  const graphPhoto: string | undefined = usePhoto(graph, photoUrl ? undefined : userId, size === 'xl' ? '240x240' : size === 'lg' ? '120x120' : '64x64');
  const photo: string | undefined = photoUrl || graphPhoto;
  const sizeClass: string = size === 'sm' ? styles.avatarSm : size === 'lg' ? styles.avatarLg : size === 'xl' ? styles.avatarXl : '';
  return (
    <span className={`${styles.avatar} ${sizeClass}`} style={{ background: photo ? undefined : colorFor(name) }} aria-hidden="true">
      {photo ? <img src={photo} alt="" /> : initials(name)}
      {presence && <span className={`${styles.presence} ${presenceClass(presence)}`} title={presence} />}
    </span>
  );
};
