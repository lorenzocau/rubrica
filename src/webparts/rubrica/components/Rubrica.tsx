import * as React from 'react';
import styles from './Rubrica.module.scss';
import type { IRubricaProps } from './IRubricaProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Rubrica extends React.Component<IRubricaProps> {
  public render(): React.ReactElement<IRubricaProps> {
    const {
      description,
    } = this.props;

    return (<section></section>
    );
  }
}
