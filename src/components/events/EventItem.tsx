import * as React from 'react';
import styles from './Events.module.scss';
import Event from '../../models/Event';
import * as moment from 'moment';

export interface IEventItemProps {
    event: Event;
    color: string;
    onClick: Function;
}

export default class EventItem extends React.Component<IEventItemProps, {}> {
    constructor(props: IEventItemProps){
        super(props);

        this._onClick = this._onClick.bind(this);
    }
    public render(): React.ReactElement<IEventItemProps> {
        const eventDate = moment(this.props.event.EventDate).format('YYYY-MM-DD');
        const endDate = moment(this.props.event.EndDate).format('YYYY-MM-DD');

        return (<div className={styles.item} onClick={this._onClick}>
                    {
                        moment(eventDate).isSame(endDate) ?
                            <div className={styles.date} style={{ backgroundColor: this.props.color }}>
                                <div className={styles.month}>{moment(this.props.event.EventDate).format('MMM').toLocaleUpperCase()}</div>
                                <div className={styles.day}>{moment(this.props.event.EventDate).format('D')}</div>
                            </div>
                        :
                            <div className={styles.date} style={{ backgroundColor: this.props.color }}>
                                <div className={styles.month}>{moment(this.props.event.EventDate).format('MMM D').toLocaleUpperCase()}</div>
                                <div className={styles.divider}></div>
                                <div className={styles.secondMonth}>{moment(this.props.event.EndDate).format('MMM D').toLocaleUpperCase()}</div>
                            </div>
                    }
                    
                    <div className={styles.text}>
                        <div className={styles.title}>{this.props.event.title.length > 40 ? `${this.props.event.title.substr(0, 40)}...` : this.props.event.title}</div>
                        {
                            this.props.event.category &&
                                <div className={styles.category}>{this.props.event.category}</div>
                        }
                        <div className={styles.time}>{moment(this.props.event.EventDate).format('ddd, MMM D, h:mm A')}</div>
                    </div>
                </div>);
    }

    private _onClick(){
        this.props.onClick(this.props.event);
    }
}