/* eslint-disable no-case-declarations */
import UserProfile from "../models/UserProfile";
import { sp } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import DateTimeHelper from "../helpers/DateTimeHelper";
import * as moment from "moment";

export interface IEmployeeDetailsService {
    getNewcomers(max: number): Promise<Array<UserProfile>>;
    getAnniversaries(max: number, range: string): Promise<Array<UserProfile>>;
}

export default class EmployeeDetailsService implements IEmployeeDetailsService {
    private _listTitle: string;
    private _ctx: WebPartContext;
    private _select: string = 'SPD_Employee/FirstName, SPD_Employee/LastName, SPD_Employee/JobTitle, SPD_Employee/EMail, SPD_HireDate, SPD_InternalHireDate';
    private _expand: string = 'SPD_Employee';

    constructor(ctx: WebPartContext, listTitle: string) {
        this._listTitle = listTitle;
        this._ctx = ctx;
    }

    public getNewcomers(max: number): Promise<Array<UserProfile>> {
        return this._getEmployeeDetails(max, '', false);
    }

    public getAnniversaries(max: number, range: string): Promise<Array<UserProfile>> {
        let filter: string;

        switch (range) {
            case 'Day':
                const date = new Date(1970, new Date().getMonth(), new Date().getDate());
                filter = ` SPD_InternalHireDate eq '${date.toISOString()}' `;
                break;
            case 'Week':
                const wStart = moment().startOf('week').toDate();
                wStart.setFullYear(1970);
                const wEnd = moment().endOf('week').toDate();
                wEnd.setFullYear(1970);
                filter = ` SPD_InternalHireDate ge '${wStart.toISOString()}' and SPD_InternalHireDate le '${wEnd.toISOString()}' `;

                break;
            default:
                const mStart = moment().startOf('month').toDate();
                mStart.setFullYear(1970);
                const mEnd = moment().endOf('month').toDate();
                mEnd.setFullYear(1970);
                filter = ` SPD_InternalHireDate ge '${mStart.toISOString()}' and SPD_InternalHireDate le '${mEnd.toISOString()}'`;
                break;
        }

        return this._getEmployeeDetails(max, filter, true);
    }

    private _getEmployeeDetails(max: number, filter: string, ascen: boolean): Promise<Array<UserProfile>> {
        return sp.web.regionalSettings.timeZone.usingCaching().get().then((rSettings) => {
            const timeZoneHours = DateTimeHelper.getSiteTimeZoneHours(rSettings);
            return sp.web.lists.getByTitle(this._listTitle).items.select(this._select).expand(this._expand).filter(filter).top(max).orderBy('SPD_HireDate', ascen).get().then((items: any) => {
                let result = new Array<UserProfile>();

                items.forEach(item => {
                    const email = item['SPD_Employee']['EMail'];
                    const user: UserProfile = {
                        email: email,
                        photoUrl: `${this._ctx.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${email}`,
                        name: `${item['SPD_Employee']['FirstName']} ${item['SPD_Employee']['LastName']}`,
                        function: item['SPD_Employee']['JobTitle'],
                        hireDate: DateTimeHelper.createDate(item['SPD_HireDate'], timeZoneHours),
                        internalHireDate: item['SPD_InternalHireDate']
                    };

                    result.push(user);
                });

                return result;
            });
        });
    }
}