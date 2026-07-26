import { mergeStyleSets } from '@fluentui/react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

/**
 * All colours come from the host site's theme (Fluent theme tokens), so the
 * web part re-themes automatically with the site design.
 */
export function getClassNames(theme?: IReadonlyTheme) { // eslint-disable-line @typescript-eslint/explicit-function-return-type
  const palette = theme && theme.palette ? theme.palette : undefined;
  const semantic = theme && theme.semanticColors ? theme.semanticColors : undefined;
  const primary = (palette && palette.themePrimary) || '#0078d4';
  const primaryDarker = (palette && palette.themeDark) || '#005a9e';
  const primarySofter = (palette && palette.themeLighterAlt) || '#eff6fc';
  const text = (semantic && semantic.bodyText) || '#323130';
  const textSub = (semantic && semantic.bodySubtext) || '#605e5c';
  const border = (semantic && semantic.bodyDivider) || '#edebe9';
  const cardBg = (semantic && semantic.bodyBackground) || '#ffffff';

  const cardShadow = '0 1.6px 3.6px rgba(0,0,0,.132), 0 0.3px 0.9px rgba(0,0,0,.108)';

  return mergeStyleSets({
    root: {
      fontSize: 14,
      color: text
    },
    hero: {
      background: `linear-gradient(135deg, ${primaryDarker}, ${primary})`,
      padding: '24px 0 0'
    },
    heroInner: {
      maxWidth: 1320,
      margin: '0 auto',
      padding: '0 24px'
    },
    searchRow: {
      position: 'relative',
      paddingBottom: 22
    },
    pivotBar: {
      background: cardBg,
      borderBottom: `1px solid ${border}`,
      selectors: {
        '.ms-Pivot': { maxWidth: 1320, margin: '0 auto', padding: '0 16px' }
      }
    },
    statusLine: {
      display: 'flex',
      alignItems: 'center',
      gap: 14,
      flexWrap: 'wrap',
      color: textSub,
      fontSize: 13,
      margin: '14px 0'
    },
    layout: {
      display: 'flex',
      gap: 20,
      alignItems: 'flex-start',
      maxWidth: 1320,
      margin: '0 auto',
      padding: '4px 24px 48px'
    },
    layoutRight: {
      flexDirection: 'row-reverse'
    },
    refiners: {
      width: 250,
      flexShrink: 0,
      selectors: {
        '@media(max-width: 760px)': { display: 'none' }
      }
    },
    refinerCard: {
      background: cardBg,
      border: `1px solid ${border}`,
      borderRadius: 8,
      boxShadow: cardShadow,
      padding: '14px 16px 10px',
      marginBottom: 12
    },
    results: {
      flex: 1,
      minWidth: 0
    },
    card: {
      background: cardBg,
      border: `1px solid ${border}`,
      borderRadius: 8,
      boxShadow: cardShadow,
      overflow: 'hidden'
    },
    rowHover: {
      display: 'flex',
      gap: 14,
      padding: '13px 16px',
      cursor: 'pointer',
      selectors: {
        ':hover': { background: primarySofter },
        ':focus': { outline: `2px solid ${primary}`, outlineOffset: -2 }
      }
    },
    fileIcon: {
      width: 38,
      height: 38,
      borderRadius: 6,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontSize: 18,
      color: '#ffffff',
      flexShrink: 0
    },
    summary: {
      fontSize: 13,
      color: textSub,
      lineHeight: 1.45,
      margin: 0,
      selectors: {
        c0: { fontWeight: 700, fontStyle: 'normal', color: text }
      }
    },
    meta: {
      fontSize: 12,
      color: textSub,
      display: 'flex',
      gap: 6,
      flexWrap: 'wrap',
      alignItems: 'center'
    },
    section: {
      marginBottom: 26
    },
    sectionHeader: {
      display: 'flex',
      alignItems: 'center',
      gap: 8,
      margin: '0 0 10px',
      fontSize: 15,
      fontWeight: 600,
      selectors: {
        i: { color: primary }
      }
    },
    peopleGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))',
      gap: 12
    },
    personCard: {
      background: cardBg,
      border: `1px solid ${border}`,
      borderRadius: 8,
      boxShadow: cardShadow,
      padding: '16px 14px 12px',
      cursor: 'pointer',
      selectors: {
        ':hover': { boxShadow: '0 6.4px 14.4px rgba(0,0,0,.132)' }
      }
    },
    siteGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(290px, 1fr))',
      gap: 12
    },
    emptyState: {
      background: cardBg,
      border: `1px dashed ${border}`,
      borderRadius: 8,
      padding: '44px 20px',
      textAlign: 'center',
      color: textSub
    },
    answerCard: {
      background: cardBg,
      border: `1px solid ${border}`,
      borderLeft: `4px solid ${primary}`,
      borderRadius: 8,
      boxShadow: cardShadow,
      padding: '16px 18px',
      marginBottom: 18
    },
    chipRow: {
      display: 'flex',
      gap: 7,
      flexWrap: 'wrap',
      marginBottom: 14
    }
  });
}

export type IUnifiedSearchClassNames = ReturnType<typeof getClassNames>;

export const FILE_TYPE_COLORS: { [type: string]: string } = {
  docx: '#2b579a', doc: '#2b579a',
  xlsx: '#217346', xls: '#217346', csv: '#217346',
  pptx: '#d24726', ppt: '#d24726',
  pdf: '#c0392b',
  one: '#7719aa',
  aspx: '#038387'
};

export const FILE_TYPE_ICONS: { [type: string]: string } = {
  docx: 'WordDocument', doc: 'WordDocument',
  xlsx: 'ExcelDocument', xls: 'ExcelDocument', csv: 'ExcelDocument',
  pptx: 'PowerPointDocument', ppt: 'PowerPointDocument',
  pdf: 'PDF',
  one: 'OneNoteLogo',
  aspx: 'FileHTML'
};

export function fileColor(type: string): string {
  return FILE_TYPE_COLORS[type] || '#605e5c';
}

export function fileIconName(type: string): string {
  return FILE_TYPE_ICONS[type] || 'Page';
}
