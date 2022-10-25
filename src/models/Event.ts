export default class Event{
    public id:number;
    public ID?:number;
    public title: string;
    public Description?: any;
    public Location?:string;
    public EventDate: Date;
    public EndDate: Date;
    public fAllDayEvent?: boolean;
    public category?: string;
    public place?: number;
    public EventType?: string;
    public RecurrenceData?: string;
    public fRecurrence?:string | boolean;
    public Duration?: number;
    public RecurrenceID?: string;
    public MasterSeriesItemID?: string;
    public url?: string;
}
