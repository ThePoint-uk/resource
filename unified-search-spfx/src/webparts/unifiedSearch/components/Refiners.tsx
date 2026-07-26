import * as React from 'react';
import { Checkbox, Text, Link, ChoiceGroup, IChoiceGroupOption, Toggle } from '@fluentui/react';
import { IRefiner, VerticalKey } from '../models';
import { IUnifiedSearchClassNames } from './styles';

export interface IMailFilters {
  dateKey: 'any' | 'day' | 'week' | 'month';
  attachmentsOnly: boolean;
}

export interface ICalendarFilters {
  onlineOnly: boolean;
}

export interface IFacetBucket {
  value: string;
  count: number;
}

export interface IRefinersProps {
  classNames: IUnifiedSearchClassNames;
  vertical: VerticalKey;

  fileRefiners?: IRefiner[];
  selectedFileTokens: string[];
  onToggleFileToken: (token: string) => void;
  onClearFileTokens: () => void;

  mailFilters: IMailFilters;
  onMailFiltersChange: (filters: IMailFilters) => void;

  calendarFilters: ICalendarFilters;
  onCalendarFiltersChange: (filters: ICalendarFilters) => void;

  departmentFacets: IFacetBucket[];
  officeFacets: IFacetBucket[];
  selectedDepartments: string[];
  selectedOffices: string[];
  onToggleDepartment: (department: string) => void;
  onToggleOffice: (office: string) => void;
}

const dateOptions: IChoiceGroupOption[] = [
  { key: 'any', text: 'Any time' },
  { key: 'day', text: 'Past 24 hours' },
  { key: 'week', text: 'Past week' },
  { key: 'month', text: 'Past month' }
];

export const Refiners: React.FunctionComponent<IRefinersProps> = props => {
  const { classNames, vertical } = props;

  const card = (title: string, children: React.ReactNode, onClear?: () => void): React.ReactNode => (
    <div className={classNames.refinerCard} key={title}>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: 8 }}>
        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>{title}</Text>
        {onClear && <Link styles={{ root: { marginLeft: 'auto', fontSize: 11 } }} onClick={onClear}>Clear</Link>}
      </div>
      {children}
    </div>
  );

  if (vertical === 'files' || vertical === 'policies' || vertical === 'all') {
    const refiners = props.fileRefiners || [];
    if (refiners.length === 0) {
      return <>{card('Filters', <Text variant="small">Refiners appear once a search returns file results.</Text>)}</>;
    }
    return (
      <>
        {refiners.map(refiner =>
          card(
            refiner.label,
            refiner.buckets.map(bucket =>
              <Checkbox
                key={bucket.filterToken}
                label={`${bucket.label} (${bucket.count})`}
                checked={props.selectedFileTokens.indexOf(bucket.filterToken) !== -1}
                onChange={() => props.onToggleFileToken(bucket.filterToken)}
                styles={{ root: { padding: '4px 0' } }}
              />),
            props.selectedFileTokens.length > 0 ? props.onClearFileTokens : undefined
          ))}
      </>
    );
  }

  if (vertical === 'mail') {
    return (
      <>
        {card('Received',
          <ChoiceGroup
            options={dateOptions}
            selectedKey={props.mailFilters.dateKey}
            onChange={(_, option) => option && props.onMailFiltersChange({
              ...props.mailFilters,
              dateKey: option.key as IMailFilters['dateKey']
            })}
          />)}
        {card('Attachments',
          <Toggle
            label="Has attachments"
            inlineLabel
            checked={props.mailFilters.attachmentsOnly}
            onChange={(_, checked) => props.onMailFiltersChange({ ...props.mailFilters, attachmentsOnly: !!checked })}
          />)}
      </>
    );
  }

  if (vertical === 'calendar') {
    return (
      <>
        {card('Meeting type',
          <Toggle
            label="Online meetings only"
            inlineLabel
            checked={props.calendarFilters.onlineOnly}
            onChange={(_, checked) => props.onCalendarFiltersChange({ onlineOnly: !!checked })}
          />)}
      </>
    );
  }

  if (vertical === 'people') {
    return (
      <>
        {props.departmentFacets.length > 0 && card('Department',
          props.departmentFacets.map(facet =>
            <Checkbox
              key={facet.value}
              label={`${facet.value} (${facet.count})`}
              checked={props.selectedDepartments.indexOf(facet.value) !== -1}
              onChange={() => props.onToggleDepartment(facet.value)}
              styles={{ root: { padding: '4px 0' } }}
            />))}
        {props.officeFacets.length > 0 && card('Office',
          props.officeFacets.map(facet =>
            <Checkbox
              key={facet.value}
              label={`${facet.value} (${facet.count})`}
              checked={props.selectedOffices.indexOf(facet.value) !== -1}
              onChange={() => props.onToggleOffice(facet.value)}
              styles={{ root: { padding: '4px 0' } }}
            />))}
        {props.departmentFacets.length === 0 && props.officeFacets.length === 0 &&
          card('Filters', <Text variant="small">Department and office filters appear with people results.</Text>)}
      </>
    );
  }

  return (
    <>{card('Filters', <Text variant="small">Site results are ranked by relevance and your activity.</Text>)}</>
  );
};
