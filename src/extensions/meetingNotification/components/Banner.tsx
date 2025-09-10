import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import styles from './Banner.module.scss';

export interface IBannerProps {
  message: string;
  onDismiss: () => void;
}

export default function Banner(props: IBannerProps) {
  const { message, onDismiss } = props;

  return (
    <div className={styles.placeholder}>
      <MessageBar
        messageBarType={MessageBarType.info}
        isMultiline={false}
        onDismiss={onDismiss}
        dismissButtonAriaLabel="Close"
      >
        {message}
      </MessageBar>
    </div>
  );
}
