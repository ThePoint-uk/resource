import * as React from 'react';
import type { IThePointMySiteProps } from './IThePointMySiteProps';
import { DashboardProvider } from '../context/DashboardContext';
import { Dashboard } from './dashboard/Dashboard';

const ThePointMySite: React.FC<IThePointMySiteProps> = (props) => (
  <DashboardProvider
    context={props.context}
    graph={props.graph}
    sp={props.sp}
    theme={props.theme}
    isDarkTheme={props.isDarkTheme}
    props={props.properties}
    isPageEditMode={props.isPageEditMode}
    instanceId={props.instanceId}
    onSaveAsDefault={props.onSaveAsDefault}
  >
    <Dashboard />
  </DashboardProvider>
);

export default ThePointMySite;
