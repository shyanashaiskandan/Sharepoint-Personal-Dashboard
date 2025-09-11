import * as React from 'react';
import { useEffect, useMemo, useState, useCallback } from 'react';
import {
  CommandBar,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  MessageBar,
  MessageBarType,
  SelectionMode,
  Stack,
  ICommandBarItemProps,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { ToDoAppProps } from './ToDoAppProps';

export type TaskStatus = 'Not Started' | 'In Progress' | 'Done';

export interface Task {
  id: number;
  name: string;         
  description?: string; 
  status: TaskStatus | string; 
}

const LIST_TITLE = 'Testing List';

export default function ToDoApp(props: ToDoAppProps) {
  const [items, setItems] = useState<Task[]>([]);
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState<{ type: MessageBarType; text: string } | null>(null);

  const columns: IColumn[] = useMemo(
    () => [
      { key: 'colName', name: 'Name', fieldName: 'name', minWidth: 160, isResizable: true },
      { key: 'colDesc', name: 'Description', fieldName: 'description', minWidth: 220, isResizable: true },
      { key: 'colStatus', name: 'Status', fieldName: 'status', minWidth: 120, maxWidth: 160, isResizable: true }
    ],
    []
  );

  const loadItems = useCallback(async () => {
    try {
      setLoading(true);
      setToast(null);

      const webUrl = props.context.pageContext.web.absoluteUrl;

      const url =
        `${webUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')/items?` +
        `$select=Id,Title,Description,Status&$orderby=Id desc`;

      const res = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) throw new Error(`SharePoint returned ${res.status}`);

      const data: { value: any[] } = await res.json();

      const mapped: Task[] = (data.value || []).map(row => ({
        id: row.Id as number,
        name: String(row.Title ?? ''),              
        description: row.Description ? String(row.Description) : '',
        status: row.Status ? String(row.Status) : ''
      }));

      setItems(mapped);
      setToast({ type: MessageBarType.success, text: `Loaded ${mapped.length} item(s) from “${LIST_TITLE}”.` });
    } catch (e: any) {
      setToast({ type: MessageBarType.error, text: `Failed to load items: ${e?.message ?? e}` });
    } finally {
      setLoading(false);
    }
  }, [props.context]);

  useEffect(() => {
    void loadItems();
  }, [loadItems]);

  const commandItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => { void loadItems(); } 
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      {toast && (
        <MessageBar messageBarType={toast.type} onDismiss={() => setToast(null)} isMultiline={false}>
          {toast.text}
        </MessageBar>
      )}

      <CommandBar items={commandItems} />

      {loading ? (
        <Spinner size={SpinnerSize.medium} label="Loading items…" />
      ) : (
        <DetailsList
          items={items}
          columns={columns}
          setKey="tasks"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      )}
    </Stack>
  );
}
