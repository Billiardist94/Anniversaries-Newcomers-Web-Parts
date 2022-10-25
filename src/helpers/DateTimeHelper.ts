import * as moment from 'moment';

export default class DateTimeHelper{
    public static getSiteTimeZoneHours(rsettings: any): number {
        let currentDateTimeOffSet: number = new Date().getTimezoneOffset() / 60;
        let siteTimeZoneBias = rsettings.Information.Bias;
        let siteTimeZoneDaylightBias = rsettings.Information.DaylightBias;

        return siteTimeZoneBias >= 0 ? ((siteTimeZoneBias / 60) - currentDateTimeOffSet) + siteTimeZoneDaylightBias / 60 : ((siteTimeZoneBias / 60) - currentDateTimeOffSet);
    }

    public static createDate(date: any, timeZoneHours: number): Date {
        return new Date(moment(date).subtract((timeZoneHours), 'hour').toISOString());
    }
}