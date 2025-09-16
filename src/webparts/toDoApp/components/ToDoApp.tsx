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
  SpinnerSize,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
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
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState<{ type: MessageBarType; text: string } | null>(null);

  const [newName, setNewName] = useState<string>('');
  const [newDescription, setNewDescription] = useState<string>('');
  const [newStatus, setNewStatus] = useState<TaskStatus>('Not Started');

  const [entityType, setEntityType] = useState<string | null>(null);

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

  const loadEntityType = useCallback(async () => {
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')?$select=ListItemEntityTypeFullName`;

      const res = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) throw new Error(`Get entity type failed (${res.status})`);

      const data = await res.json();
      setEntityType(data.ListItemEntityTypeFullName as string);
    } catch (e: any) {
      setToast({ type: MessageBarType.error, text: `Failed to read list metadata: ${e?.message ?? e}` });
    }
  }, [props.context]);

  useEffect(() => {
    void loadEntityType(); 
    void loadItems();      
  }, [loadEntityType, loadItems]);

  const addItem = useCallback(async () => {
    if (!newName.trim()) {
      setToast({ type: MessageBarType.warning, text: 'Please enter a task name before adding.' });
      return;
    }
    if (!entityType) {
      setToast({ type: MessageBarType.warning, text: 'Preparing list metadata. Please try again.' });
      return;
    }

    try {
      setSaving(true);
      setToast(null);

      const webUrl = props.context.pageContext.web.absoluteUrl;
      const url = `${webUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')/items`;

      const body = JSON.stringify({
        __metadata: { type: entityType },           
        Title: newName.trim(),                       
        Description: newDescription.trim(),          
        Status: newStatus                            
      });

      const res: SPHttpClientResponse = await props.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': '3.0'
          },
          body
        }
      );

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Create failed (${res.status}). ${text}`);
      }

      setNewName('');
      setNewDescription('');
      setNewStatus('Not Started');
      setToast({ type: MessageBarType.success, text: 'Task added successfully.' });
      await loadItems();
    } catch (e: any) {
      setToast({ type: MessageBarType.error, text: `Failed to add item: ${e?.message ?? e}` });
    } finally {
      setSaving(false);
    }
  }, [entityType, newName, newDescription, newStatus, props.context, loadItems]);

  const commandItems: ICommandBarItemProps[] = [
    { key: 'refresh', text: 'Refresh', iconProps: { iconName: 'Refresh' }, onClick: () => { void loadItems(); } }
  ];

  const statusOptions: IDropdownOption[] = [
    { key: 'Not Started', text: 'Not Started' },
    { key: 'In Progress', text: 'In Progress' },
    { key: 'Done', text: 'Done' }
  ];

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      {toast && (
        <MessageBar messageBarType={toast.type} onDismiss={() => setToast(null)} isMultiline={false}>
          {toast.text}
        </MessageBar>
      )}

      <CommandBar items={commandItems} />

      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { maxWidth: 640 } }}>
        <TextField
          label="Task name"
          placeholder="e.g., Set up SPFx project"
          value={newName}
          onChange={(_, v) => setNewName(v ?? '')}
          required
        />
        <TextField
          label="Description"
          placeholder="Optional details…"
          multiline
          value={newDescription}
          onChange={(_, v) => setNewDescription(v ?? '')}
        />
        <Dropdown
          label="Status"
          selectedKey={newStatus}
          options={statusOptions}
          onChange={(_, option) => option && setNewStatus(option.key as TaskStatus)}
        />
        <PrimaryButton
          text={saving ? 'Adding…' : 'Add task'}
          onClick={() => void addItem()}
          disabled={saving || !newName.trim()}
        />
      </Stack>

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
