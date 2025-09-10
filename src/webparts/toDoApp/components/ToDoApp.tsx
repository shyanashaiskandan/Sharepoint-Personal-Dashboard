// src/webparts/todo/components/ToDoApp.tsx
import * as React from 'react';
import { useState, useMemo } from 'react';
import {
  CommandBar,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Panel,
  PanelType,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  SelectionMode,
  Stack,
  ICommandBarItemProps
} from '@fluentui/react';

import { ToDoAppProps } from './ToDoAppProps';

export type TaskStatus = 'Not Started' | 'In Progress' | 'Done';

export interface Task {
  id: string;
  name: string;
  description?: string;
  status: TaskStatus;
}

const STATUS_OPTIONS: TaskStatus[] = ['Not Started', 'In Progress', 'Done'];

function newId() {
  return Math.random().toString(36).slice(2);
}

const initialMock: Task[] = [
  { id: newId(), name: 'Task 1', description: 'Description 1', status: 'Done' },
  { id: newId(), name: 'Task 2', description: 'Description 2', status: 'In Progress' },
  { id: newId(), name: 'Task 3', description: 'Description 3', status: 'Not Started' }
];

export default function ToDoApp(_props: ToDoAppProps) {

  const [items, setItems] = useState<Task[]>(initialMock);
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [formName, setFormName] = useState('');
  const [formDesc, setFormDesc] = useState('');
  const [formStatus, setFormStatus] = useState<TaskStatus>('Not Started');
  const [toast, setToast] = useState<{ type: MessageBarType; text: string } | null>(null);

  const columns: IColumn[] = useMemo(
    () => [
      { key: 'colName', name: 'Name', fieldName: 'name', minWidth: 160, isResizable: true },
      { key: 'colDesc', name: 'Description', fieldName: 'description', minWidth: 220, isResizable: true },
      { key: 'colStatus', name: 'Status', fieldName: 'status', minWidth: 120, maxWidth: 140, isResizable: true }
    ],
    []
  );

  const statusOptions: IDropdownOption[] = STATUS_OPTIONS.map(s => ({ key: s, text: s }));

  const onAddClick = () => {
    setFormName('');
    setFormDesc('');
    setFormStatus('Not Started');
    setIsAddOpen(true);
  };

  const onSave = () => {
    if (!formName.trim()) {
      setToast({ type: MessageBarType.error, text: 'Name is required.' });
      return;
    }
    const newTask: Task = {
      id: newId(),
      name: formName.trim(),
      description: formDesc.trim() || '',
      status: formStatus
    };
    setItems(prev => [newTask, ...prev]);
    setIsAddOpen(false);
    setToast({ type: MessageBarType.success, text: 'Task added (local only for now).' });
  };

  const commandItems: ICommandBarItemProps[] = [
    { key: 'add', text: 'Add Task', iconProps: { iconName: 'Add' }, onClick: onAddClick },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => setToast({ type: MessageBarType.info, text: 'Refresh will load from SharePoint later.' })
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

      <DetailsList
        items={items}
        columns={columns}
        setKey="tasks"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
      />

      <Panel
        isOpen={isAddOpen}
        type={PanelType.medium}
        onDismiss={() => setIsAddOpen(false)}
        headerText="Add Task"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField label="Name" value={formName} onChange={(_, v) => setFormName(v || '')} required />
          <TextField
            label="Description"
            value={formDesc}
            onChange={(_, v) => setFormDesc(v || '')}
            multiline
            rows={3}
          />
          <Dropdown
            label="Status"
            options={statusOptions}
            selectedKey={formStatus}
            onChange={(_, opt) => opt && setFormStatus(opt.key as TaskStatus)}
          />

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Save" onClick={onSave} />
            <DefaultButton text="Cancel" onClick={() => setIsAddOpen(false)} />
          </Stack>
        </Stack>
      </Panel>
    </Stack>
  );
}
